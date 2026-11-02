"""Descarga el Excel de últimos precios de grifos (EVPC) de Osinergmin y arma la página SuperFuel.

Uso:
  python actualizar.py               -> descarga el Excel de Osinergmin
  python actualizar.py archivo.xlsx  -> usa un Excel ya descargado
Crea la carpeta _site/ con index.html y los íconos.
"""
import datetime as dt
import io
import json
import shutil
import sys
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


def descargar() -> bytes:
    hoy = dt.datetime.now(dt.timezone(dt.timedelta(hours=-5)))
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


def armar_datos(xlsx: bytes) -> dict:
    df = pd.read_excel(io.BytesIO(xlsx))
    lima = df[(df.DEPARTAMENTO == "LIMA") & (df.PROVINCIA == "LIMA") & (df.PRODUCTO_ACTIVO == "SI")]
    lima = lima[lima.PRODUCTO.isin(PRODUCTOS)]
    if lima.empty:
        sys.exit("El Excel no trae grifos de Lima Metropolitana; no actualizo la página.")
    distritos = sorted(lima.DISTRITO.unique())
    actividades = sorted(lima.ACTIVIDAD.astype(str).str.strip().unique())
    grifos, idx = [], {}
    precios = {k: [] for k in PRODUCTOS.values()}
    for _, r in lima.sort_values("CODIGO_OSINERG").iterrows():
        cod = int(r.CODIGO_OSINERG)
        if cod not in idx:
            idx[cod] = len(grifos)
            grifos.append([
                " ".join(str(r.RAZON).split()),
                distritos.index(r.DISTRITO),
                " ".join(str(r.DIRECCION).split()),
                actividades.index(str(r.ACTIVIDAD).strip()),
            ])
        fecha = pd.to_datetime(r.FCHA_REGISTRO)
        precios[PRODUCTOS[r.PRODUCTO]].append([idx[cod], round(float(r.PRECIO_VENTA), 2), fecha.strftime("%Y-%m-%d %H:%M")])
    snap = pd.to_datetime(lima.FCHA_REGISTRO).max().strftime("%Y-%m-%d %H:%M")
    print(f"{len(grifos)} grifos, {sum(map(len, precios.values()))} precios, último registro {snap}")
    return {"snap": snap, "d": distritos, "a": actividades, "s": grifos, "p": precios}


def main():
    xlsx = Path(sys.argv[1]).read_bytes() if len(sys.argv) > 1 else descargar()
    datos = json.dumps(armar_datos(xlsx), ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    plantilla = (AQUI / "plantilla.html").read_text(encoding="utf-8")
    sitio = AQUI / "_site"
    sitio.mkdir(exist_ok=True)
    (sitio / "index.html").write_text(plantilla.replace("__DATA__", datos), encoding="utf-8")
    for icono in ("apple-touch-icon.png", "icon-192.png"):
        shutil.copy(AQUI / icono, sitio / icono)
    print("Listo: _site/index.html")


if __name__ == "__main__":
    main()
