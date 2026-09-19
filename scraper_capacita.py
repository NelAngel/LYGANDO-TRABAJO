#!/usr/bin/env python3
"""
SCRAPER CURSOS CAPACITA-T (MTPE) - capacitacionlaboral.trabajo.gob.pe/cursos/
Extrae el catálogo de cursos gratuitos de Capacita que ofrece el Ministerio de
Trabajo y Promoción del Empleo, y guarda:
  - cursos.xlsx   -> Excel en el MISMO formato de control que ofertas.xlsx
  - cursos.json   -> respaldo en JSON

Uso:
  python3 scraper_capacita.py             # listado de la portada (solo ~12 cursos)
  python3 scraper_capacita.py --todo      # CATÁLOGO COMPLETO: lee el sitemap de WordPress (214 cursos)
  python3 scraper_capacita.py --dias MAX  # para control (aunque cursos NO caducan)

El listado /cursos/ solo muestra ~12 cursos y su ?pg=N repite la misma página.
El catálogo COMPLETO está en el sitemap: /wp-sitemap-posts-cursos-1.xml (214 fichas).

NOTA sobre vigencia: los CURSOS (a diferencia de las ofertas/convocatorias) NO
caducan -> la columna `visible` siempre queda en `si` salvo que el curso no tenga
ficha valida. `purgar.py` NO toca cursos.xlsx (ver REGLA DE VIGENCIA en AGENTS.md).
"""

import argparse
import json
import re
import sys
import time
from datetime import datetime
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

BASE_URL = "https://capacitacionlaboral.trabajo.gob.pe"
CURSOS_URL = BASE_URL + "/cursos/"
SITEMAP_URL = BASE_URL + "/wp-sitemap-posts-cursos-1.xml"
OUT_XLSX = "cursos.xlsx"
OUT_JSON = "cursos.json"
EMPRESA = "MTPE - Capacita"
FUENTE = "Capacita-T (MTPE)"

# 13 columnas de control/base (misma base que ofertas y convocatorias) + extras
# propias del scraping de cursos:
#   duracion      : horas del curso (texto "N horas")
#   certifica     : "si"/"no" (el curso entrega certificado MTPE)
#   segmento      : ENTIDAD QUE EMITE el curso (MTPE, Cisco, Huawei, Fundación
#                  Romero o ABC del BCP). Se deduce del LMS donde se accede
#                  (columna empezar_curso): la página agrupa los cursos por este
#                  segmento (de entidad emisora).
#   empezar_curso : enlace al LMS (boton "EMPEZAR CURSO" -> capacitacion.../course/view.php)
#   ver_ruta      : enlace a la ruta formativa de la que forma parte
CLAVES_ENLACE = [
    ("EMPEZAR CURSO", "empezar_curso"),
    ("IR A LA RUTA", "ver_ruta"),
    ("VER RUTA COMPLETA", "ver_ruta"),
    ("EMPEZAR", "empezar_curso"),
]

COLUMNAS = ["id", "titulo", "empresa", "ubicacion", "sueldo", "descripcion",
            "enlace", "whatsapp", "fecha", "fuente", "destacado", "visible",
            "contacto", "duracion", "certifica", "segmento", "empezar_curso",
            "ver_ruta"]

# Entidad que emite cada curso según la plataforma a la que lleva "EMPEZAR CURSO".
# El certificado de cada ficha lo confirma (Fundación Romero / MTPE / ABC del BCP);
# Cisco y Huawei no imprimen ese texto pero sus cursos se dictan en sus plataformas.
DOMINIO_SEGMENTO = {
    "capacitate.trabajo.gob.pe": "MTPE",
    "lms.becasgruporomero.pe": "Fundación Romero",
    "www.netacad.com": "Cisco",
    "e.huawei.com": "Huawei",
    "www.viabcp.com": "ABC del BCP",
}


def segmento_de(empezar_curso):
    """Segmento (entidad emisora) a partir de la URL del LMS."""
    from urllib.parse import urlparse
    dom = urlparse(empezar_curso or "").netloc
    return DOMINIO_SEGMENTO.get(dom, "Otro")

HEADERS = {
    "User-Agent": "Mozilla/5.0 (X11; Linux aarch64) AppleWebKit/537.36 "
                  "Chrome/151.0.0.0 Safari/537.36",
}

DIAS_MAX = 10


def obtener_pagina(url):
    try:
        r = requests.get(url, headers=HEADERS, timeout=30)
        r.raise_for_status()
        return BeautifulSoup(r.text, "lxml")
    except Exception as e:
        print(f"[error] {url}: {e}")
        return None


def a_plaintext(html):
    """Convierte HTML a texto plano (quita etiquetas, normaliza espacios)."""
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", html or "")).strip()


def extraer_campos_ficha(soup):
    """De la ficha de un curso extrae titulo, horas, certifica y los enlaces clave."""
    f = {"titulo": "", "duracion": "", "certifica": "", "descripcion": "",
         "empezar_curso": "", "ver_ruta": ""}

    h1 = soup.find("h1")
    if h1:
        f["titulo"] = h1.get_text(" ", strip=True)

    # Duracion: patron "N horas" en el texto de la ficha.
    texto = soup.get_text(" ", strip=True)
    m = re.search(r"(\d+)\s*horas?", texto, re.IGNORECASE)
    if m:
        f["duracion"] = m.group(0)

    # Certifica: buscar "certificad" (MTPE/MTP).
    if re.search(r"certificad[oa]?", texto, re.IGNORECASE):
        f["certifica"] = "si"

    # Descripcion: primer parrafo con masa de texto util, saltando el menu,
    # el breadcrumb y la barra de botones (Empezar curso / Ver Ruta...).
    RUIDO = ("empezar curso", "ver ruta", "ingresa", "registrate", "recuperar",
             "rutas formativas", "plataforma ministerio", "cursos de la ruta",
             "navegación", "inicio ")
    contenedor = soup.select_one(".columna-b") or soup
    for tag in ("p", "li", "div"):
        for et in contenedor.find_all(tag):
            t = et.get_text(" ", strip=True)
            bajo = t.lower()
            if 60 <= len(t) <= 1200 and not any(k in bajo for k in RUIDO):
                f["descripcion"] = t
                break
        if f["descripcion"]:
            break

    # Enlaces del LMS y ruta formativa.
    for a in soup.find_all("a", href=True):
        etiqueta = a.get_text(" ", strip=True)
        for clave, campo in CLAVES_ENLACE:
            if clave in etiqueta.upper() and not f[campo]:
                f[campo] = urljoin(BASE_URL, a["href"])
    return f


def listar_cursos(soup):
    """De la pagina /cursos/ (o /cursos/page/N/) extrae [titulo, url] de cada curso."""
    cursos = []
    for a in soup.find_all("a", href=True):
        href = a.get("href", "")
        url = urljoin(BASE_URL, href)
        # Fichas individuales: /cursos/<slug>/ (no el listado, ni paginas)
        if not re.match(r"^" + re.escape(CURSOS_URL) + r"[^/]+/$", url):
            continue
        titulo = a.get_text(" ", strip=True)
        if not titulo:
            continue
        if not any(c[0] == titulo and c[1] == url for c in cursos):
            cursos.append((titulo, url))
    return cursos


def obtener_de_sitemap():
    """CATALOGO COMPLETO de cursos desde el sitemap de WordPress.
    El listado /cursos/ solo muestra ~12, pero el sitemap trae las 214 fichas."""
    try:
        r = requests.get(SITEMAP_URL, headers=HEADERS, timeout=30)
        r.raise_for_status()
    except Exception as e:
        print(f"[error] sitemap: {e}")
        return []
    urls = []
    for loc in re.findall(r"<loc>([^<]+)</loc>", r.text):
        if re.match(r"^" + re.escape(CURSOS_URL) + r"[^/]+/$", loc) and loc not in urls:
            urls.append(loc)
    return urls


def urls_del_listado():
    """URLs de la portada /cursos/ (solo las ~12 que muestra el listado)."""
    sopa = obtener_pagina(CURSOS_URL)
    if sopa is None:
        return []
    urls = []
    for _, u in listar_cursos(sopa):
        if u not in urls:
            urls.append(u)
    return urls


def guardar_xlsx(filas):
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Cursos"
    relleno = PatternFill("solid", fgColor="27ae60")
    fuente_cab = Font(color="FFFFFF", bold=True)
    for j, nombre in enumerate(COLUMNAS, 1):
        c = ws.cell(row=1, column=j, value=nombre)
        c.fill = relleno
        c.font = fuente_cab
    for i, fila in enumerate(filas, 2):
        for j, nombre in enumerate(COLUMNAS, 1):
            ws.cell(row=i, column=j, value=fila.get(nombre, ""))
    anchos = [5, 45, 16, 22, 10, 90, 32, 12, 12, 18, 10, 8, 12, 12, 10, 20, 36, 36]
    for j, ancho in enumerate(anchos, 1):
        ws.column_dimensions[get_column_letter(j)].width = ancho
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLUMNAS))}{len(filas) + 1}"
    wb.save(OUT_XLSX)
    print(f"[ok] Excel guardado: {OUT_XLSX} ({len(filas)} filas)")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pag", type=int, default=1, help="pagina del listado por la que empezar")
    ap.add_argument("--todo", action="store_true", help="recorrer TODA la paginacion")
    ap.add_argument("--verificar", action="store_true", help="no usado; mantener CLI estable")
    args = ap.parse_args()

    print("=" * 60)
    print(" SCRAPER CURSOS CAPACITA-T (MTPE)")
    print("=" * 60)

    # --todo -> CATALOGO COMPLETO desde el sitemap de WordPress (214 cursos).
    # Sin --todo -> solo el listado de portada (~12). El ?pg=N repite la misma
    # pagina, por eso el sitemap es la unica fuente del catalogo completo.
    if args.todo:
        urls = obtener_de_sitemap()
        if urls:
            print(f"[sitemap] catalogo completo: {len(urls)} cursos")
        else:
            print("[sitemap] no disponible -> uso el listado de portada")
            urls = urls_del_listado()
    else:
        urls = urls_del_listado()

    if not urls:
        print("[error] no se detectaron cursos")
        sys.exit(1)

    print(f"\nCursos detectados: {len(urls)}")

    filas = []
    for i, url in enumerate(urls, 1):
        ficha = obtener_pagina(url)
        if ficha is None:
            continue
        campos = extraer_campos_ficha(ficha)
        slug = url.rstrip("/").rsplit("/", 1)[-1].replace("-", " ").strip()
        titulo_real = campos["titulo"] or slug.capitalize()
        fila = {
            "id": i,
            "titulo": titulo_real,
            "empresa": EMPRESA,
            "ubicacion": "Nivel Nacional (virtual) - Todo el país",
            "sueldo": "Gratuito",
            "descripcion": campos["descripcion"],
            "enlace": url,
            "whatsapp": "",
            "fecha": datetime.now().strftime("%Y-%m-%d"),
            "fuente": FUENTE,
            "destacado": "no",
            "visible": "si" if titulo_real else "no",   # regla de oro: sin titulo no se publica
            "contacto": "",
            "duracion": campos["duracion"],
            "certifica": campos["certifica"],
            "segmento": segmento_de(campos["empezar_curso"]),
            "empezar_curso": campos["empezar_curso"],
            "ver_ruta": campos["ver_ruta"],
        }
        filas.append(fila)
        print(f"  [{i}] {titulo_real[:60]}")
        if args.verificar and i >= 5:   # mantener el placeholder estable pero acotado
            break
        time.sleep(0.15)                # cortesía con el servidor (214 fichas)

    guardar_xlsx(filas)
    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump({
            "fuente": FUENTE,
            "url": CURSOS_URL,
            "fecha_scraping": datetime.now().isoformat(),
            "total": len(filas),
            "cursos": filas,
        }, f, ensure_ascii=False, indent=2)
    print(f"[ok] JSON guardado: {OUT_JSON}")


if __name__ == "__main__":
    main()
