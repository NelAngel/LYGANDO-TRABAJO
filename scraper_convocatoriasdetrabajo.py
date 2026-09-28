#!/usr/bin/env python3
"""Scraper de convocatorias del Estado: www.convocatoriasdetrabajo.com
(empleos y prácticas del sector público: CAS, 728, 276, prácticas, etc.)

Genera `convocatoriasdetrabajo.xlsx` (mismas 17 columnas que convocatorias.xlsx,
incluida `vigencia` = fecha límite) + respaldo `convocatoriasdetrabajo.json`.
La página (script.js) lo carga como fuente "Convocatorias" y las mantiene
visibles hasta su fecha LIMITE (regla de vigencia, igual que las ONPE).
Al terminar PURGA automáticamente las que ya vencieron (purgar_por_vigencia()).

Uso:
  python3 scraper_convocatoriasdetrabajo.py          -> TODAS las vigentes de la portada
  python3 scraper_convocatoriasdetrabajo.py --max 50 -> solo 50 detalles (pruebas)
Se detecta SOLO la lista de la portada (bloques con "Vigente hasta el DD/MM/AAAA"),
se deduplican y se salta el detalle de las que ya pasaron (estado finalizado).
"""
import sys
import json
import random
import re
import time
import hashlib
import os
from datetime import datetime, date

import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin, urlparse

from purgar import purgar_por_vigencia

BASE_URL = "https://www.convocatoriasdetrabajo.com/"
DOMINIO = "www.convocatoriasdetrabajo.com"
FUENTE = "Convocatoriasdetrabajo.com"

ARCHIVO_JSON = "convocatoriasdetrabajo.json"
ARCHIVO_XLSX = "convocatoriasdetrabajo.xlsx"

MIN_DELAY = 0.8
MAX_DELAY = 1.6
TIMEOUT = 30

# Campos que escupe el scrapeo (los mismos 17 que convocatorias.xlsx)
BLANCOS = {
    "ver_detalles": "", "funciones": "", "guia_registro": "", "postular": "",
    "whatsapp": "", "destacado": "no",
}

HEADERS = {
    "User-Agent": ("Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/151.0.0.0 Safari/537.36"),
    "Accept-Language": "es-PE,es;q=0.9,en;q=0.8",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Connection": "keep-alive",
}

session = requests.Session()
session.headers.update(HEADERS)

MESES = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
    "julio": 7, "agosto": 8, "setiembre": 9, "septiembre": 9,
    "octubre": 10, "noviembre": 11, "diciembre": 12,
}


def esperar():
    time.sleep(random.uniform(MIN_DELAY, MAX_DELAY))


def limpiar_texto(texto):
    if not texto:
        return ""
    return re.sub(r"\s+", " ", texto.replace("\xa0", " ")).strip()


def texto_elemento(el):
    return limpiar_texto(el.get_text(" ", strip=True)) if el else ""


def hash_url(url):
    return hashlib.sha256(url.encode("utf-8")).hexdigest()[:16]


def es_url_valida(url):
    try:
        p = urlparse(url)
        if p.scheme not in ("http", "https"):
            return False
        d = p.netloc.lower()
        return d == DOMINIO or d.endswith("." + DOMINIO)
    except Exception:
        return False


def descargar(url):
    try:
        esperar()
        r = session.get(url, timeout=TIMEOUT, allow_redirects=True)
        r.raise_for_status()
        return r.text
    except requests.RequestException:
        return None


# --------------------------------------------------------------- LISTADO

def extraer_activas(html, url_actual):
    """Convocatorias vigentes de la portada: bloques cuyo texto trae
    'Vigente hasta el DD/MM/AAAA'. Devuelve {url: {'vigencia':ISO,'region':''}}."""
    soup = BeautifulSoup(html, "lxml")
    hoy = datetime.now().date()
    encontradas = {}
    for a in soup.find_all("a", href=True):
        href = a.get("href") or ""
        if "oferta-de-empleo-" not in href.lower():
            continue
        url = urljoin(url_actual, href).split("#")[0]
        if not es_url_valida(url) or "?" in url:
            continue
        node = a
        texto = ""
        for _ in range(7):
            if node.parent is None:
                break
            node = node.parent
            texto = node.get_text(" ", strip=True)
            if "Vigente hasta el" in texto:
                break
        m = re.search(r"[Vv]igente hasta el\s*(\d{1,2})/(\d{1,2})/(\d{4})", texto)
        if not m:
            continue
        try:
            vig = date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except ValueError:
            continue
        if vig < hoy:
            continue  # ya venció en el propio listado
        encontradas.setdefault(url, {"vigencia": vig.isoformat(), "region": ""})
    return encontradas


def region_de_texto(texto):
    todos = [
        "amazonas", "ancash", "apurimac", "arequipa", "ayacucho", "cajamarca",
        "callao", "cusco", "huancavelica", "huanuco", "ica", "junin",
        "la libertad", "lambayeque", "lima", "loreto", "madre de dios",
        "moquegua", "pasco", "piura", "puno", "san martin", "tacna",
        "tumbes", "ucayali",
    ]
    t = limpiar_texto(texto).lower()
    for reg in todos:
        if re.search(r"\b" + reg + r"\b", t):
            return reg.title().replace("Madre De Dios", "Madre de Dios").replace("San Martin", "San Martín")
    return ""


# --------------------------------------------------------------- DETALLE

def extraer_campo(texto, etiquetas):
    for etiqueta in etiquetas:
        for patron in (
            re.escape(etiqueta) + r"\s*:\s*\n\s*([^\n]+)",
            re.escape(etiqueta) + r"\s*:\s*([^\n]+)",
        ):
            m = re.search(patron, texto, re.IGNORECASE)
            if m:
                return limpiar_texto(m.group(1))
    return ""


def fecha_larga_a_iso(texto):
    """'25 de Septiembre del 2026' -> '2026-09-25'. '' si no es fecha larga."""
    m = re.search(
        r"(\d{1,2})\s+de\s+([a-záéíóúñ]+)\s+(?:del\s+|de\s+)(\d{4})",
        texto, re.IGNORECASE,
    )
    if not m:
        return ""
    mes = MESES.get(m.group(2).lower())
    if not mes:
        return ""
    try:
        return date(int(m.group(3)), mes, int(m.group(1))).isoformat()
    except ValueError:
        return ""


def extraer_uuid(url):
    m = re.search(r"-(\d+)\.html?$", url)
    return m.group(1) if m else hash_url(url)


def scrapear_detalle(url, region_listado):
    """Extrae los campos de la convocatoria (una tarjeta por convocatoria)."""
    html = descargar(url)
    if not html:
        return None
    soup = BeautifulSoup(html, "lxml")
    full = soup.get_text("\n", strip=True)

    titulo = texto_elemento(soup.find("h1"))

    institucion = extraer_campo(full, ["Institución", "Institucion"])
    contrato = extraer_campo(full, ["Tipo de contrato"])
    nivel = extraer_campo(full, ["Hay puestos para"])
    lugar = extraer_campo(full, ["Lugar de labores"])
    remuneracion = extraer_campo(full, ["Remuneración", "Remuneracion"])
    carreras = extraer_campo(full, ["Hay plazas para"])

    # Puestos (encabezados tipo "CAS N° 015 - ITEM N° 01: MADRE SUSTITUTA")
    puestos = []
    for h in soup.find_all(["h2", "h3", "h4"]):
        t = texto_elemento(h)
        tl = t.lower()
        if (
            "item n°" in tl or "item no" in tl or "cas n°" in tl or "cas no" in tl
            or "plaza n°" in tl or "puesto" in tl
            or re.search(r"\b(especialista|analista|asistente|auxiliar|coordinador|"
                         r"técnico|tecnico|profesional|abogado|ingeniero|médico|medico|"
                         r"enfermera|obstetra|secretaria|conductor|digitador)\b", tl)
        ) and t not in puestos:
            puestos.append(t)
            if len(puestos) >= 8:
                break

    estado = ""
    if "CONVOCATORIA FINALIZADA" in full.upper():
        estado = "FINALIZADA"

    # Vigencia: fecha límite ("esta vigente hasta el DD/MM/AAAA" o "Finaliza: dd de mmmm")
    vig = ""
    m = re.search(r"[Vv]igente hasta el\s*(\d{1,2})/(\d{1,2})/(\d{4})", full)
    if m:
        try:
            vig = date(int(m.group(3)), int(m.group(2)), int(m.group(1))).isoformat()
        except ValueError:
            vig = ""
    if not vig:
        fin = extraer_campo(full, ["Finaliza", "Fecha límite"])
        vig = fecha_larga_a_iso(fin)

    # Fecha de publicación
    fecha = fecha_larga_a_iso(extraer_campo(full, ["Fecha de Publicación", "Fecha de Publicacion"]))

    # Descripción (resumen amigable de la convocatoria)
    partes = []
    if contrato:
        partes.append(f"Contrato: {contrato}")
    if nivel:
        partes.append(f"Hay puestos para: {nivel}")
    if carreras:
        partes.append(f"Plazas para: {carreras}")
    if remuneracion:
        partes.append(f"Remuneración: {remuneracion}")
    if puestos:
        partes.append("Puestos: " + "; ".join(puestos))
    descripcion = " · ".join(partes)

    if not lugar:
        lugar = region_listado
    if not lugar and re.search(r"nivel nacional|todo el país|todo el pais|todas las region", full, re.I):
        lugar = "Nivel Nacional"

    visible = "si" if (titulo and lugar and estado != "FINALIZADA") else "no"

    return {
        "id": extraer_uuid(url),
        "titulo": titulo,
        "empresa": institucion,
        "ubicacion": lugar,
        "sueldo": remuneracion,
        "descripcion": descripcion,
        "enlace": url,
        "fecha": fecha,
        "fuente": FUENTE,
        "visible": visible,
        "vigencia": vig,
        **BLANCOS,
    }


# --------------------------------------------------------------- GUARDADO

def guardar_json(datos):
    with open(ARCHIVO_JSON, "w", encoding="utf-8") as f:
        json.dump(datos, f, ensure_ascii=False, indent=2)
    print(f"[ok] JSON guardado: {ARCHIVO_JSON}")


def guardar_excel(datos):
    import openpyxl
    columnas = ["id", "titulo", "empresa", "ubicacion", "sueldo", "descripcion",
                "enlace", "whatsapp", "fecha", "fuente", "destacado", "visible",
                "ver_detalles", "funciones", "guia_registro", "postular", "vigencia"]
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Convocatorias"
    ws.append(columnas)
    for fila in datos:
        ws.append([fila.get(c, "") for c in columnas])
    anchos = [6, 45, 28, 22, 24, 70, 60, 8, 12, 26, 10, 8, 26, 26, 26, 42, 12]
    for i, w in enumerate(anchos, 1):
        ws.column_dimensions[chr(64 + i)].width = w
    wb.save(ARCHIVO_XLSX)
    print(f"[ok] Excel guardado: {ARCHIVO_XLSX} ({len(datos)} filas)")


# --------------------------------------------------------------- EJECUCIÓN

def ejecutar(max_detalles=0):
    print("=" * 70)
    print("SCRAPER CONVOCATORIASDETRABAJO.COM")
    print("=" * 70)

    html = descargar(BASE_URL)
    if not html:
        print("ERROR: no se pudo leer la portada.")
        return

    activas = extraer_activas(html, BASE_URL)
    print(f"Convocatorias vigentes en la portada: {len(activas)}")

    ordenadas = sorted(activas.items(),
                       key=lambda kv: (kv[1]["vigencia"], kv[0]), reverse=True)
    if max_detalles:
        ordenadas = ordenadas[:max_detalles]
        print(f"(procesando detalles de las {len(ordenadas)} primeras; sin --max procesa todas)")

    resultados = {}
    for pos, (url, info) in enumerate(ordenadas, 1):
        print(f"[{pos}/{len(ordenadas)}] {url.split('/')[-1][:60]}")
        detalle = scrapear_detalle(url, info["region"])
        if not detalle:
            continue
        # seguridad: si la vigencia ya pasó en el detalle, se descarta/oculta
        if detalle["vigencia"]:
            try:
                if date.fromisoformat(detalle["vigencia"]) < date.today():
                    detalle["visible"] = "no"
            except ValueError:
                pass
        resultados[url] = detalle
        if pos % 10 == 0:
            guardar_json(list(resultados.values()))

    # redondear la región de los que no traen lugar por departamento bien escrito
    finales = list(resultados.values())

    guardar_json(finales)
    guardar_excel(finales)

    borradas = purgar_por_vigencia(ARCHIVO_XLSX)
    print(f"[ok] Purga automática: {borradas} convocatorias con vigencia pasada eliminadas.")
    print(f"Total guardadas: {len(finales)}")


if __name__ == "__main__":
    max_det = 0  # por defecto: TODAS las vigentes detectadas en la portada
    for arg in sys.argv[1:]:
        if arg.startswith("--max") and sys.argv.index(arg) + 1 < len(sys.argv):
            max_det = int(sys.argv[sys.argv.index(arg) + 1])
    if any(a == "--todo" for a in sys.argv[1:]):
        max_det = 0
    ejecutar(max_det)