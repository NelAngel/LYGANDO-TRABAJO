#!/usr/bin/env python3
"""Purgar anuncios viejos de los Excels (BD de la página).

  - ofertas.xlsx (fotos)      -> se purgan por ANTIGÜEDAD (más de DIAS_MAX días).
  - convocatorias.xlsx (ONPE) -> se purgan por su fecha LÍMITE (columna `vigencia`):
    cuando la vigencia YA PASÓ se borran; las "hasta completar vacantes" (sin fecha
    límite) se quedan mientras el scraping las siga publicando.
  - convocatoriasdetrabajo.xlsx (del Estado) -> igual que las ONPE, por `vigencia`.

Uso:
  python3 purgar.py            -> limpia los Excels con su regla correspondiente
  python3 purgar.py --dias 10  -> usa una cantidad distinta de días para las FOTOS

También es importable: `from purgar import purgar, purgar_por_vigencia, DIAS_MAX`
(los scrapers llaman solo a purgar_por_vigencia al terminar).
"""
import os
import sys
from datetime import date, timedelta

import openpyxl

DIAS_MAX = 10
ARCHIVOS = ["ofertas.xlsx", "convocatorias.xlsx"]
CONV_ESTADO = "convocatoriasdetrabajo.xlsx"


def serie_a_iso(v):
    if v is None:
        return None
    if isinstance(v, (int, float)) and v > 20000:
        d = date(1899, 12, 30) + timedelta(days=int(v))
        return d.isoformat()
    return str(v).strip()


def purgar(archivo, dias=DIAS_MAX, silencio=False):
    wb = openpyxl.load_workbook(archivo)
    ws = wb.active
    corte = date.today() - timedelta(days=dias)
    borradas = 0
    for fila in range(ws.max_row, 1, -1):
        fecha_raw = ws.cell(row=fila, column=9).value  # columna 'fecha'
        iso = serie_a_iso(fecha_raw)
        if iso is None:
            continue
        try:
            f = date.fromisoformat(iso)
        except ValueError:
            continue
        if f < corte:
            ws.delete_rows(fila, 1)
            borradas += 1
    wb.save(archivo)
    if not silencio:
        print(f"  {archivo}: borradas {borradas} filas (corte {corte}, anterior a {dias} días)")
    return borradas


def purgar_por_vigencia(archivo, silencio=False):
    """Borra las filas cuya columna `vigencia` (fecha límite) YA PASÓ.
    Las filas sin fecha límite ("hasta completar vacantes") se conservan."""
    wb = openpyxl.load_workbook(archivo)
    ws = wb.active
    indice = None
    for j in range(1, ws.max_column + 1):
        if str(ws.cell(row=1, column=j).value).strip().lower() == "vigencia":
            indice = j
            break
    if indice is None:
        wb.close()
        if not silencio:
            print(f"  {archivo}: sin columna 'vigencia' (aún sin scraping), no se purga.")
        return 0
    hoy = date.today()
    borradas = 0
    for fila in range(ws.max_row, 1, -1):
        raw = ws.cell(row=fila, column=indice).value
        iso = serie_a_iso(raw)
        if not iso:
            continue
        try:
            if date.fromisoformat(iso) < hoy:
                ws.delete_rows(fila, 1)
                borradas += 1
        except ValueError:
            continue
    wb.save(archivo)
    if not silencio:
        print(f"  {archivo}: borradas {borradas} filas (vigencia pasada, desde {hoy})")
    return borradas


def purgar_todo(dias=DIAS_MAX):
    """Fotos: por antigüedad; ONPE y convocatorias del Estado: por vigencia."""
    total = purgar(ARCHIVOS[0], dias)
    total += purgar_por_vigencia(ARCHIVOS[1])
    if os.path.exists(CONV_ESTADO):
        total += purgar_por_vigencia(CONV_ESTADO)
    return total


if __name__ == "__main__":
    dias = DIAS_MAX
    for arg in sys.argv[1:]:
        if arg.startswith("--dias") and sys.argv.index(arg) + 1 < len(sys.argv):
            dias = int(sys.argv[sys.argv.index(arg) + 1])
    print(f"Purgando fotos con más de {dias} días y convocatorias con vigencia pasada...")
    purgar_todo(dias)
    print("Listo.")