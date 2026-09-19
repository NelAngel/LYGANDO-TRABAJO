#!/usr/bin/env python3
"""
SCRAPER CONVOCATORIAS ONPE (portaltrabajos.pe)
Lee el feed JSON de Blogger de la etiqueta ONPE y guarda:
  - convocatorias.xlsx   -> Excel con el MISMO formato que ofertas.xlsx (la página lo lee)
  - convocatorias_onpe.json -> respaldo en JSON

Uso:
  python3 scraper_onpe.py            # ultimas 150 convocatorias
  python3 scraper_onpe.py --max 50   # ultimas 50
  python3 scraper_onpe.py --todo     # todas (767+), pagina el feed
"""

import argparse
import json
import re
import sys
import time
from datetime import datetime

import requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

# Purga automática de anuncios viejos (regla de 10 días): se ejecuta al guardar.
from purgar import DIAS_MAX, purgar

FEED_URL = "https://www.portaltrabajos.pe/feeds/posts/default/-/ONPE"
OUT_XLSX = "convocatorias.xlsx"
OUT_JSON = "convocatorias_onpe.json"
EMPRESA = "ONPE"
FUENTE = "Portaltrabajo - ONPE"
# 12 columnas originales + 4 enlaces clave que vienen en cada entrada del feed:
#   ver_detalles : PDF "[ VER MÁS DETALLES ]"   (bases / requisitos completos)
#   funciones    : PDF "[ FUNCIONES ]"
#   guia_registro: PDF "[ VER GUÍA DE REGISTRO ]"
#   postular     : enlace al sistema de ONPE "[ POSTULAR ]"
COLUMNAS = ["id", "titulo", "empresa", "ubicacion", "sueldo", "descripcion",
            "enlace", "whatsapp", "fecha", "fuente", "destacado", "visible",
            "ver_detalles", "funciones", "guia_registro", "postular"]

HEADERS = {
    "User-Agent": "Mozilla/5.0 (X11; Linux aarch64) AppleWebKit/537.36 "
                  "Chrome/151.0.0.0 Safari/537.36"
}


def descargar(url, params=None):
    try:
        r = requests.get(url, params=params, headers=HEADERS, timeout=25)
        r.raise_for_status()
        return r.json()
    except Exception as e:
        print(f"[error] {url} {e}")
        return None


def a_plaintext(html):
    texto = re.sub(r"<[^>]+>", " ", html or "")
    texto = texto.replace("&nbsp;", " ").replace("&amp;", "&")
    texto = re.sub(r"\s+", " ", texto)
    return texto.strip()


ETIQUETAS = [
    "Instituci\u00f3n", "Vacantes", "Ubicaci\u00f3n", "Fecha de Publicaci\u00f3n",
    "Vigente", "Salario", "Remuneraci\u00f3n", "PUESTO", "NIVEL", "TIPO", "CONVOCATORIA",
    "N\u00famero de plazas", "Requisitos", "FECHA", "Inscripciones",
]


def extraer_campo(texto, clave):
    """Valor tras 'clave: ...' hasta la siguiente etiqueta conocida."""
    m = re.search(
        rf"{re.escape(clave)}\s*[:\-]\s*(.*?)(?=\s+(?:{'|'.join(ETIQUETAS)})\s*[:\-]|\Z)",
        texto, re.IGNORECASE | re.DOTALL)
    return a_plaintext(m.group(1)[:160]) if m else ""


def limpiar_ubicacion(txt):
    """'Nivel Nacional - (Según ODPE Disponible)' -> 'Nivel Nacional'."""
    return re.sub(r"\s*[-]\s*\(\s*[Ss]eg[uú]n\s+ODPE[^)]*\)|\s*\(\s*[Ss]eg[uú]n\s+ODPE[^)]*\)", "", txt).strip()


# Enlaces que interesan de cada entrada y a qué columna van.
# Se leen del HTML crudo del feed (el <a href="..."> con su texto).
ETIQUETAS_ENLACE = {
    "VER MÁS DETALLES": "ver_detalles",
    "FUNCIONES": "funciones",
    "VER GUÍA DE REGISTRO": "guia_registro",
    "POSTULAR": "postular",
    "VER CONVOCATORIAS Y POSTULAR": "postular",
    "VER BASES": "ver_detalles",
}


def extraer_enlaces(html):
    """Busca los anchors '[ VER MÁS DETALLES ]', '[ FUNCIONES ]', '[ POSTULAR ]'...
    y devuelve {columna: href} con la primera ocurrencia de cada uno."""
    res = {}
    for m in re.finditer(r'<a\s+[^>]*href=["\']([^"\']+)["\'][^>]*>(.*?)</a>',
                         html or "", re.IGNORECASE | re.DOTALL):
        href, txt = m.groups()
        txt = a_plaintext(txt).upper()
        clave = txt.replace(" ", "")
        for nombre, columna in ETIQUETAS_ENLACE.items():
            if clave == nombre.replace(" ", "") or nombre.replace(" ", "") in clave:
                res.setdefault(columna, href.strip())
                break
    return res


def convertir(entradas):
    """Convierte las entries del feed al formato de la hoja."""
    filas = []
    for i, e in enumerate(entradas, 1):
        titulo = a_plaintext(e.get("title", {}).get("$t", ""))
        publicado = e.get("published", {}).get("$t", "")
        enlace = next((l["href"] for l in e.get("link", []) if l.get("rel") == "alternate"), "")
        html = e.get("content", {}).get("$t", "")
        contenido = a_plaintext(html)

        ubicacion = limpiar_ubicacion(extraer_campo(contenido, "Ubicaci\u00f3n"))
        fila = {
            "id": i,
            "titulo": titulo,
            "empresa": EMPRESA,
            "ubicacion": ubicacion,
            "sueldo": extraer_campo(contenido, "Salario"),
            "descripcion": contenido[:330],
            "enlace": enlace,
            "whatsapp": "",
            "fecha": publicado[:10],
            "fuente": FUENTE,
            "destacado": "no",
            # Regla de oro: sin LUGAR no se publica (queda en el Excel, oculta)
            "visible": "si" if ubicacion else "no",
        }
        fila.update(extraer_enlaces(html))
        filas.append(fila)
    return filas


def enlace_valido(f):
    """Señales fuertes: status 200, URL en portaltrabajos.pe y página con contenido.
    (Ojo: este template de Blogger sirve <data:...> sin renderizar incluso en posts
    vivos, por eso NO se usa como señal de página muerta.)"""
    try:
        r = requests.get(f["enlace"], headers=HEADERS, timeout=15)
        if r.status_code != 200:
            return False
        if "portaltrabajos.pe" not in r.url:
            return False
        if len(r.text) < 5000:
            return False
        return True
    except Exception:
        return False


def verificar_enlaces(filas):
    """Comprueba que cada enlace responda 200, sin redirección rara y con el título."""
    print(f"\nVerificando {len(filas)} enlaces (puede tardar)...")
    malos = 0
    for f in filas:
        if not enlace_valido(f):
            malos += 1
            f["visible"] = "no"
            print("  [inválido]", f["titulo"][:60])
        time.sleep(0.4)
    print(f"Verificación: {len(filas) - malos}/{len(filas)} enlaces válidos, {malos} ocultos")
    return filas


def guardar_xlsx(filas):
    wb = Workbook()
    ws = wb.active
    ws.title = "Ofertas"
    encabezado_fill = PatternFill("solid", fgColor="E95420")
    encabezado_font = Font(color="FFFFFF", bold=True)

    for j, nombre in enumerate(COLUMNAS, 1):
        c = ws.cell(row=1, column=j, value=nombre)
        c.fill = encabezado_fill
        c.font = encabezado_font
        c.alignment = Alignment_h()

    for i, fila in enumerate(filas, 2):
        for j, nombre in enumerate(COLUMNAS, 1):
            ws.cell(row=i, column=j, value=fila.get(nombre, ""))

    anchos = [6, 52, 12, 22, 24, 70, 70, 12, 12, 24, 10, 8, 48, 48, 48, 42]
    for j, ancho in enumerate(anchos, 1):
        ws.column_dimensions[get_column_letter(j)].width = ancho
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLUMNAS))}{len(filas)+1}"

    wb.save(OUT_XLSX)
    print(f"[ok] Excel guardado: {OUT_XLSX} ({len(filas)} filas)")


def Alignment_h():
    from openpyxl.styles import Alignment
    return Alignment(horizontal="left", vertical="center")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--max", type=int, default=150, help="ultimas N (default 150)")
    ap.add_argument("--todo", action="store_true", help="traer todas paginando el feed")
    ap.add_argument("--verificar", action="store_true",
                    help="comprobar cada enlace (HTTP 200 + título) y ocultar los caídos")
    args = ap.parse_args()

    print("=" * 64)
    print("  SCRAPER CONVOCATORIAS ONPE - portaltrabajos.pe")
    print("=" * 64)

    dato = descargar(FEED_URL, {"alt": "json", "max-results": 150, "start-index": 1})
    if not dato:
        print("[error] no se pudo leer el feed")
        sys.exit(1)

    total = int(dato["feed"].get("openSearch$totalResults", {}).get("$t", 0))
    print(f"Convocatorias ONPE disponibles en el sitio: {total}")

    entradas = list(dato["feed"].get("entry", []))
    if args.todo:
        inicio = 151
        while inicio <= total:
            time.sleep(1)
            b = descargar(FEED_URL, {"alt": "json", "max-results": 150, "start-index": inicio})
            if not b:
                break
            entradas.extend(b["feed"].get("entry", []))
            inicio += 150
    else:
        entradas = entradas[:args.max]

    print(f"Procesadas {len(entradas)} convocatorias (recientes)...")
    filas = convertir(entradas)

    if args.verificar:
        filas = verificar_enlaces(filas)

    guardar_xlsx(filas)
    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump({
            "fuente": "https://www.portaltrabajos.pe/search/label/ONPE",
            "institucion": EMPRESA,
            "fecha_scraping": datetime.now().isoformat(),
            "total": len(filas),
            "convocatorias": filas,
        }, f, ensure_ascii=False, indent=2)
    print(f"[ok] JSON guardado: {OUT_JSON}")

    # Purga automática (regla de 10 días): deja solo las convocatorias recientes.
    borradas = purgar(OUT_XLSX)
    print(f"[ok] Purga automática: {borradas} convocatorias antiguas eliminadas"
          f" (solo quedan las de {DIAS_MAX} días).")


if __name__ == "__main__":
    main()