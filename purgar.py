#!/usr/bin/env python3
"""Purgar anuncios con más de DIAS_MAX días de antigüedad de ofertas.xlsx y convocatorias.xlsx.

Uso:
  python3 purgar.py            -> borra filas con fecha < hoy - DIAS_MAX
  python3 purgar.py --dias 10  -> usa una cantidad distinta de días

También es importable: `from purgar import purgar, DIAS_MAX` (el scraper lo llama solo).
"""
import sys
from datetime import date, timedelta

import openpyxl

DIAS_MAX = 10
ARCHIVOS = ["ofertas.xlsx", "convocatorias.xlsx"]


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


def purgar_todo(dias=DIAS_MAX):
    """Limpia todos los Excels. Devuelve el total de filas borradas."""
    total = 0
    for a in ARCHIVOS:
        total += purgar(a, dias)
    return total


if __name__ == "__main__":
    dias = DIAS_MAX
    for arg in sys.argv[1:]:
        if arg.startswith("--dias") and sys.argv.index(arg) + 1 < len(sys.argv):
            dias = int(sys.argv[sys.argv.index(arg) + 1])
    print(f"Purgando anuncios con más de {dias} días...")
    purgar_todo(dias)
    print("Listo.")