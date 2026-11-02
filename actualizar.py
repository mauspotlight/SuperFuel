"""Descarga el Excel de últimos precios de grifos (EVPC) de Osinergmin y arma la página SuperFuel.

Uso:
  python actualizar.py               -> descarga el Excel de Osinergmin
  python actualizar.py archivo.xlsx  -> usa un Excel ya descargado
  SIN_UBICAR=1 python actualizar.py  -> no busca ubicaciones nuevas (solo usa coords.json)

Crea la carpeta _site/ con index.html y los íconos, y actualiza coords.json
con la ubicación (latitud, longitud) de cada grifo.
"""
import datetime as dt
import io
import json
import math
import os
import re
import shutil
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

import pandas as pd

BASE = "https://www.osinergmin.gob.pe/seccion/centro_documental/hidrocarburos/SCOP/SCOP-DOCS/{year}/Registro-precios/Ultimos-Precios-Registrados-EVPC.xlsx"
PRODUCTOS = {
    "GASOHOL REGULAR": "GR",
    "GASOHOL PREMIUM": "GP",
    "Diesel B5 S-50 UV": "DB5",
    "GLP - G": "GLP",
    "GAS NATURAL VEHICULAR COMPRIMIDO": "GNV",
    "GAS NATURAL VEHICULAR LICUEFACTADO": "GNL",
}
AQUI = Path(__file__).parent
COORDS = AQUI / "coords.json"
UA = "SuperFuel/1.0 (+https://github.com/mauspotlight/SuperFuel)"
LIMA = (-12.55, -77.25, -11.70, -76.60)  # sur, oeste, norte, este
MAX_BUSQUEDAS = 2500     # tope de consultas a OpenStreetMap por ejecución
REINTENTAR_DIAS = 30     # vuelve a buscar los grifos no encontrados cada 30 días
PEGAR_A_GRIFO_M = 400    # si hay un grifo de OpenStreetMap a menos de esto, usa su posición exacta
LIMA_TZ = dt.timezone(dt.timedelta(hours=-5))


# ---------- descarga del Excel ----------
def descargar() -> bytes:
    hoy = dt.datetime.now(LIMA_TZ)
    errores = []
    for year in (hoy.year, hoy.year - 1):  # en enero el archivo puede seguir en la carpeta del año anterior
        url = BASE.format(year=year)
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (SuperFuel)"})
            with urllib.request.urlopen(req, timeout=120) as r:
                data = r.read()
            if data[:2] == b"PK":  # un .xlsx real empieza así
                print(f"Descargado {url} ({len(data):,} bytes)")
                return data
            errores.append(f"{url}: no es un Excel")
        except Exception as e:  # noqa: BLE001
            errores.append(f"{url}: {e}")
    sys.exit("No se pudo descargar el Excel:\n" + "\n".join(errores))


# ---------- ubicaciones ----------
def get_json(url: str, data: bytes | None = None):
    req = urllib.request.Request(url, data=data, headers={"User-Agent": UA, "Accept-Language": "es"})
    with urllib.request.urlopen(req, timeout=180) as r:
        return json.loads(r.read().decode("utf-8"))


def grifos_osm() -> list[tuple[float, float]]:
    """Todos los grifos (amenity=fuel) que OpenStreetMap tiene en Lima."""
    s, w, n, e = LIMA
    q = f'[out:json][timeout:120];(node["amenity"="fuel"]({s},{w},{n},{e});way["amenity"="fuel"]({s},{w},{n},{e}););out center;'
    for url in ("https://overpass-api.de/api/interpreter", "https://overpass.kumi.systems/api/interpreter"):
        try:
            res = get_json(url, urllib.parse.urlencode({"data": q}).encode())
            pts = []
            for el in res.get("elements", []):
                lat = el.get("lat") or el.get("center", {}).get("lat")
                lon = el.get("lon") or el.get("center", {}).get("lon")
                if lat and lon:
                    pts.append((lat, lon))
            print(f"OpenStreetMap: {len(pts)} grifos en Lima")
            return pts
        except Exception as e:  # noqa: BLE001
            print(f"Overpass no respondió ({url}): {e}")
    return []


def distancia_m(a, b):
    r = math.pi / 180
    x = (b[0] - a[0]) * r
    y = (b[1] - a[1]) * r
    h = math.sin(x / 2) ** 2 + math.cos(a[0] * r) * math.cos(b[0] * r) * math.sin(y / 2) ** 2
    return 2 * 6371000 * math.asin(math.sqrt(h))


ABREV = [
    (r"\bAV(DA)?\.?\s", "AVENIDA "), (r"\bJR\.?\s", "JIRON "), (r"\bCA\.?\s", "CALLE "),
    (r"\bPROL(ONG)?\.?\s", "PROLONGACION "), (r"\bPJE\.?\s", "PASAJE "), (r"\bCARR?\.?\s", "CARRETERA "),
    (r"\bPANAM\.?\s", "PANAMERICANA "), (r"\bSTA\.?\s", "SANTA "), (r"\bSTO\.?\s", "SANTO "),
]
CORTES = r"\s(ESQ\.?|ESQUINA|CON\s|INTERSEC|INT\.|MZ\.?|MZA\.?|MANZANA|LOTE|LT\.?|URB\.?|URBANIZACION|SECTOR|FRENTE|ALT\.|ALTURA|A\.?\s?H\.?|ASOC\.?|ASOCIACION|COOP\.?|P\.?\s?J\.?|PUEBLO JOVEN|ETAPA|PARCELA|SUB\s?LOTE|S/N|Y|CDRA\.?|CUADRA)\b"


def candidatos(direccion: str) -> list[str]:
    s = direccion.upper()
    s = re.sub(r"\(.*?\)", " ", s)
    s = re.sub(r"\b(N[°ºO]\.?|NRO\.?|NUM\.?|NO\.)\s*", " ", s)
    s = s.replace("°", " ").replace("º", " ")
    for a, b in ABREV:
        s = re.sub(a, b, s)
    s = re.sub(r"\s+", " ", s).strip(" ,.-")
    partes = [p.strip(" ,.-") for p in s.split(",")]
    # si la dirección empieza con un pueblo joven, urbanización, etc., usa la parte que nombra la vía
    via = next((p for p in partes if re.match(r"(AVENIDA|JIRON|CALLE|PROLONGACION|PASAJE|CARRETERA|PANAMERICANA)\b", p)), partes[0])
    corto = re.split(CORTES, " " + via)[0].strip(" ,.-")
    sin_num = re.sub(r"\s\d[\d\-/ ]*$", "", corto).strip(" ,.-")
    out = []
    for c in (corto, via, sin_num, partes[0]):
        c = c.strip()
        if len(c) >= 6 and c not in out:
            out.append(c)
    return out


class Nominatim:
    def __init__(self):
        self.usadas = 0
        self.bloqueado = False
        self.ultimo = 0.0

    def buscar(self, texto: str, distrito: str):
        if self.bloqueado or self.usadas >= MAX_BUSQUEDAS:
            return None
        espera = 1.1 - (time.time() - self.ultimo)  # política de uso: máximo 1 consulta por segundo
        if espera > 0:
            time.sleep(espera)
        s, w, n, e = LIMA
        params = {
            "q": f"{texto}, {distrito.title()}, Lima, Peru",
            "format": "jsonv2", "limit": 1, "countrycodes": "pe",
            "viewbox": f"{w},{n},{e},{s}", "bounded": 1,
        }
        self.usadas += 1
        self.ultimo = time.time()
        try:
            res = get_json("https://nominatim.openstreetmap.org/search?" + urllib.parse.urlencode(params))
        except urllib.error.HTTPError as err:
            if err.code in (403, 429):
                print(f"OpenStreetMap pidió parar ({err.code}); sigo mañana.")
                self.bloqueado = True
            return None
        except Exception:  # noqa: BLE001
            return None
        if res:
            return float(res[0]["lat"]), float(res[0]["lon"])
        return None


def ubicar(grifos: dict[int, tuple[str, str]], buscar: bool) -> dict[str, list | None]:
    """grifos: {codigo: (direccion, distrito)} -> {codigo: [lat, lon] o None}"""
    cache = json.loads(COORDS.read_text(encoding="utf-8")) if COORDS.exists() else {}
    hoy = dt.datetime.now(LIMA_TZ).date()
    pendientes = []
    for cod, (direccion, distrito) in grifos.items():
        c = cache.get(str(cod))
        if c and c.get("dir") == direccion:
            if c.get("ll") or (hoy - dt.date.fromisoformat(c.get("t", "2000-01-01"))).days < REINTENTAR_DIAS:
                continue
        pendientes.append(cod)
    print(f"Ubicaciones: {len(grifos) - len(pendientes)} ya conocidas, {len(pendientes)} por buscar")

    if buscar and pendientes:
        osm = grifos_osm()
        nom = Nominatim()
        hechos = 0
        for cod in pendientes:
            if nom.bloqueado or nom.usadas >= MAX_BUSQUEDAS:
                break
            direccion, distrito = grifos[cod]
            ll = None
            for texto in candidatos(direccion):
                ll = nom.buscar(texto, distrito)
                if ll or nom.bloqueado:
                    break
            if nom.bloqueado:
                break
            fuente = "dir"
            if ll and osm:
                cerca = min(osm, key=lambda p: distancia_m(ll, p))
                if distancia_m(ll, cerca) <= PEGAR_A_GRIFO_M:
                    ll, fuente = cerca, "osm"
            cache[str(cod)] = {
                "dir": direccion,
                "ll": [round(ll[0], 6), round(ll[1], 6)] if ll else None,
                "src": fuente if ll else None,
                "t": hoy.isoformat(),
            }
            hechos += 1
            if hechos % 50 == 0:
                COORDS.write_text(json.dumps(cache, ensure_ascii=False, indent=0, sort_keys=True), encoding="utf-8")
                print(f"  {hechos}/{len(pendientes)} buscados")
        COORDS.write_text(json.dumps(cache, ensure_ascii=False, indent=0, sort_keys=True), encoding="utf-8")
        print(f"Buscados {hechos} grifos con {nom.usadas} consultas")

    out = {}
    for cod in grifos:
        c = cache.get(str(cod))
        out[str(cod)] = c.get("ll") if c else None
    con = sum(1 for v in out.values() if v)
    print(f"Grifos con ubicación: {con} de {len(out)}")
    return out


# ---------- datos de la página ----------
def armar_datos(xlsx: bytes, buscar: bool) -> dict:
    df = pd.read_excel(io.BytesIO(xlsx))
    lima = df[(df.DEPARTAMENTO == "LIMA") & (df.PROVINCIA == "LIMA") & (df.PRODUCTO_ACTIVO == "SI")]
    lima = lima[lima.PRODUCTO.isin(PRODUCTOS)]
    if lima.empty:
        sys.exit("El Excel no trae grifos de Lima Metropolitana; no actualizo la página.")
    distritos = sorted(lima.DISTRITO.unique())
    actividades = sorted(lima.ACTIVIDAD.astype(str).str.strip().unique())

    info = {}
    for _, r in lima.iterrows():
        info.setdefault(int(r.CODIGO_OSINERG), (" ".join(str(r.DIRECCION).split()), str(r.DISTRITO)))
    coords = ubicar(info, buscar)

    grifos, idx = [], {}
    precios = {k: [] for k in PRODUCTOS.values()}
    for _, r in lima.sort_values("CODIGO_OSINERG").iterrows():
        cod = int(r.CODIGO_OSINERG)
        if cod not in idx:
            idx[cod] = len(grifos)
            ll = coords.get(str(cod))
            grifos.append([
                " ".join(str(r.RAZON).split()),
                distritos.index(r.DISTRITO),
                " ".join(str(r.DIRECCION).split()),
                actividades.index(str(r.ACTIVIDAD).strip()),
                round(ll[0], 5) if ll else None,
                round(ll[1], 5) if ll else None,
            ])
        fecha = pd.to_datetime(r.FCHA_REGISTRO)
        precios[PRODUCTOS[r.PRODUCTO]].append([idx[cod], round(float(r.PRECIO_VENTA), 2), fecha.strftime("%Y-%m-%d %H:%M")])
    snap = pd.to_datetime(lima.FCHA_REGISTRO).max().strftime("%Y-%m-%d %H:%M")
    print(f"{len(grifos)} grifos, {sum(map(len, precios.values()))} precios, último registro {snap}")
    return {"snap": snap, "d": distritos, "a": actividades, "s": grifos, "p": precios}


def main():
    xlsx = Path(sys.argv[1]).read_bytes() if len(sys.argv) > 1 else descargar()
    buscar = os.environ.get("SIN_UBICAR") != "1"
    datos = json.dumps(armar_datos(xlsx, buscar), ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    plantilla = (AQUI / "plantilla.html").read_text(encoding="utf-8")
    sitio = AQUI / "_site"
    sitio.mkdir(exist_ok=True)
    (sitio / "index.html").write_text(plantilla.replace("__DATA__", datos), encoding="utf-8")
    for icono in ("apple-touch-icon.png", "icon-192.png"):
        shutil.copy(AQUI / icono, sitio / icono)
    print("Listo: _site/index.html")


if __name__ == "__main__":
    main()
