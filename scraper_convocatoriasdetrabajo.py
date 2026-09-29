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
from datetime import datetime, date, timedelta

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
MESES_CORTOS = {
    "ene": 1, "feb": 2, "mar": 3, "abr": 4, "may": 5, "jun": 6,
    "jul": 7, "ago": 8, "sep": 9, "sept": 9, "oct": 10, "nov": 11, "dic": 12,
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

# La portada marca la vigencia en cada tarjeta con un pill relativo:
#   "Finaliza en N días"  / "Finaliza en 1 días" / "Finaliza hoy"
# (antes usaba "Vigente hasta el DD/MM/AAAA"; el sitio lo cambió 2026-09-28).
# Se mantiene el patrón absoluto como fallback por si lo recuperan.
PAT_FINALIZA_EN = re.compile(r"[Ff]inaliza\s+(?:en\s+)?(\d{1,3})\s+d[ií]as?")
PAT_HASTA_EL = re.compile(r"[Vv]igente hasta el\s*(\d{1,2})/(\d{1,2})/(\d{4})")


def dias_de_pill(art, hoy):
    """Días de vigencia que anuncia la tarjeta (0 = hoy). None si no hay pill."""
    pill = texto_elemento(art.find("span", class_=re.compile("main__pills")))
    m = PAT_FINALIZA_EN.search(pill)
    if m:
        try:
            return int(m.group(1))
        except ValueError:
            return None
    if "finaliza hoy" in pill.lower():
        return 0
    # compat: texto absoluto dentro de la tarjeta
    m2 = PAT_HASTA_EL.search(pill + " " + art.get_text(" ", strip=True))
    if m2:
        try:
            return (date(int(m2.group(3)), int(m2.group(2)), int(m2.group(1))) - hoy).days
        except ValueError:
            return None
    return None


def normalizar_clave(s):
    return re.sub(r"[^a-z0-9]+", "", limpiar_texto(s).lower())


def extraer_activas(html, url_actual, slug="oferta-de-empleo-"):
    """Convocatorias vigentes de una página: cada `article.convocatoria` con su
    pill "Finaliza en N días". En la portada el slug es `oferta-de-empleo-`
    (una tarjeta por convocatoria); en los listados por CARRERA es
    `oportunidad-laboral-` (una tarjeta por PUESTO/código CAS).
    Devuelve {url: {'vigencia':ISO,'region':''}}."""
    soup = BeautifulSoup(html, "lxml")
    hoy = datetime.now().date()
    encontradas = {}
    for art in soup.find_all("article", class_="convocatoria"):
        dias = dias_de_pill(art, hoy)
        if dias is None:
            continue
        for a in art.find_all("a", href=True):
            href = a.get("href") or ""
            if slug not in href.lower():
                continue
            url = urljoin(url_actual, href).split("#")[0]
            if not es_url_valida(url) or "?" in url:
                continue
            if dias < 0:
                continue  # la tarjeta ya quedó atrasada respecto a hoy
            vig = hoy + timedelta(days=dias)
            encontradas.setdefault(url, {"vigencia": vig.isoformat(), "region": ""})
            break  # una tarjeta = una convocatoria/puesto
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


def extraer_datos_clave(soup):
    """Campos del nuevo panel "Datos clave" del detalle: cada `div.dato` tiene
    `.dato__label > span` (etiqueta) y `.dato__value` (valor) SIN dos puntos
    (el sitio rediseñó la ficha 2026-09-28 y rompió extraer_campo). Devuelve
    {etiqueta_normalizada: valor} p.ej. {'institucion': 'SENASA'}."""
    datos = {}
    for bloque in soup.select("div.dato"):
        label = bloque.select_one(".dato__label span") or bloque.select_one(".dato__label")
        valor = bloque.select_one(".dato__value")
        if label is None:
            continue
        k = normalizar_clave(texto_elemento(label))
        if not k:
            continue
        v = texto_elemento(valor) if valor else ""
        if k not in datos:
            datos[k] = v
    return datos


# Las páginas por PUESTO (`oportunidad-laboral-...`) NO traen el panel
# "Datos clave": usan el bloque "Claves del puesto" con el VALOR ENCIMA de la
# etiqueta (mismo div padre): p.ej. "Ucayali\nLugar de labores".
CLAVES_PUESTO_LABELS = {
    "lugardelabores": ["Lugar de labores", "Lugar de trabajo"],
    "remuneracion": ["Remuneración", "Remuneracion"],
    "plazoparapostular": ["Plazo para postular", "Finaliza"],
    "vacantedisponible": ["Vacante disponible", "Vacantes"],
}
def extraer_claves_puesto(soup):
    datos = {}
    for k, etiquetas in CLAVES_PUESTO_LABELS.items():
        for etq in etiquetas:
            for el in soup.find_all(string=lambda t: " ".join(t.split()).strip() == etq):
                print_ok = el.parent
                hijo = print_ok.parent
                if hijo is None:
                    continue
                cadenas = [texto_elemento(p) for p in hijo.children]
                if etq not in cadenas:
                    continue
                i = cadenas.index(etq)
                # el valor está ANTES del label, pero puede haber divs vacíos
                # de por medio: tomar el último texto NO vacío anterior.
                for j in range(i - 1, -1, -1):
                    if cadenas[j]:
                        datos[k] = cadenas[j]
                        break
                if k in datos:
                    break
            if k in datos:
                break
    return datos


def extraer_breadcrumb(full):
    """'INEI > Convocatoria OPERADOR DE ...' -> 'INEI'. Complementa a
    extraer_datos_clave para las páginas por puesto (no tienen div.dato)."""
    m = re.search(r"(?:^|\n|>)\s*([A-ZÁÉÍÓÚÑ0-9&.'´()\- ]{1,45}?)\s*>\s*Convocatoria\b",
                  full[:1200], re.IGNORECASE)
    if not m:
        return ""
    instit = m.group(1).strip()
    if instit.lower() in ("inicio", "home", "convocatorias"):
        return ""
    return instit


def fecha_corta_a_iso(texto):
    """'25 sept 2026' -> '2026-09-25'. '' si no es ese formato."""
    m = re.search(r"\b(\d{1,2})\s+([a-záéíóúñ]{3,5})\s+(\d{4})\b", texto, re.IGNORECASE)
    if not m:
        return ""
    mes = MESES_CORTOS.get(m.group(2).lower())
    if not mes:
        return ""
    try:
        return date(int(m.group(3)), mes, int(m.group(1))).isoformat()
    except ValueError:
        return ""


def fecha_rango_a_iso(texto):
    """CRONOGRAMA: 'Del 14 de setiembre al 05 de octubre del 2026' ->
    ('2026-09-14', '2026-10-05'). Acepta 'Del 14 al 25 de setiembre del 2026' y
    'del 14/09/2026 al 05/10/2026'. Es el INICIO REAL de postulación (lo pidió el
    usuario 2026-09-28: poner "fecha de inicio y cuándo termina" en los trabajos
    con enlace)."""
    def mes_num(txt):
        t = limpiar_texto(txt).lower().rstrip(".")
        return MESES.get(t) or MESES_CORTOS.get(t)

    # 'Del 14 de setiembre al 05 de octubre del 2026' (el año va al final)
    m = re.search(
        r"del\s+(\d{1,2})\s+de\s+([a-záéíóúñ.]{3,10}).{0,40}?"
        r"al\s+(\d{1,2})\s+de\s+([a-záéíóúñ.]{3,10})\s+(?:de\s+|del\s+)?(\d{4})",
        texto, re.IGNORECASE | re.DOTALL)
    if m:
        d1, mes1, d2, mes2, anio = m.groups()
        m1, m2 = mes_num(mes1), mes_num(mes2)
        if m1 and m2:
            try:
                return (date(int(anio), m1, int(d1)).isoformat(),
                        date(int(anio), m2, int(d2)).isoformat())
            except ValueError:
                return ("", "")
    # 'Del 14 al 25 de setiembre del 2026'
    m = re.search(
        r"del\s+(\d{1,2})\s+al\s+(\d{1,2})\s+de\s+([a-záéíóúñ.]{3,10})\s+(?:de\s+|del\s+)?(\d{4})",
        texto, re.IGNORECASE)
    if m:
        d1, d2, mes, anio = m.groups()
        mm = mes_num(mes)
        if mm:
            try:
                return (date(int(anio), mm, int(d1)).isoformat(),
                        date(int(anio), mm, int(d2)).isoformat())
            except ValueError:
                return ("", "")
    # 'del 14/09/2026 al 05/10/2026'
    m = re.search(r"(\d{1,2})/(\d{1,2})/(\d{4})\s+al\s+(\d{1,2})/(\d{1,2})/(\d{4})", texto, re.I)
    if m:
        d1, m1, a1, d2, m2, a2 = m.groups()
        return (f"{int(a1):04d}-{int(m1):02d}-{int(d1):02d}",
                f"{int(a2):04d}-{int(m2):02d}-{int(d2):02d}")
    return ("", "")


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


# Enlaces directos que la propia convocatoria ofrece en su ficha (NO llevar a
# convocatoriasdetrabajo.com sino directo al trámite):
#   postular : botón "POSTULA AQUÍ" / "INSCRÍBETE" / "APLICA" -> portal de la
#              institución (ej. aplicativo.pj.gob.pe, reclutamiento.onpe.gob.pe).
#   bases    : "Ver aquí Bases (convocatoria completa y cronograma)", anexos...
#              Normalmente PDF en Google Drive o repositorio de la institución.
#   video    : "(VIDEO) Cómo postular al ..." -> video tutorial (YouTube).
# Se ignoran los enlaces internos del sitio y el texto genérico del navegador.
PAT_VIDEO    = re.compile(r"video|tutorial", re.IGNORECASE)
PAT_POSTULAR = re.compile(r"postul|inscr|aplica|registra|regís|ingresa", re.IGNORECASE)
PAT_BASE     = re.compile(r"base|cronograma|convocatoria completa|anexo|declaraci", re.IGNORECASE)


def extraer_enlaces(soup):
    directos = {"postular": "", "bases": "", "video": ""}
    for a in soup.find_all("a", href=True):
        href = (a.get("href") or "").strip()
        if not href.lower().startswith(("http://", "https://")):
            continue
        if DOMINIO in href.lower():
            continue  # enlaces internos del propio sitio, no sirven
        texto = texto_elemento(a)
        if len(texto) < 2 or len(texto) > 90:
            continue
        if PAT_VIDEO.search(texto):
            if not directos["video"]:
                directos["video"] = href
        elif PAT_POSTULAR.search(texto):
            if not directos["postular"]:
                directos["postular"] = href
        elif PAT_BASE.search(texto):
            if not directos["bases"]:
                directos["bases"] = href
    return directos


def scrapear_detalle(url, region_listado, categoria=""):
    """Extrae los campos de la convocatoria (una tarjeta por convocatoria)."""
    html = descargar(url)
    if not html:
        return None
    soup = BeautifulSoup(html, "lxml")
    full = soup.get_text("\n", strip=True)
    clave = extraer_datos_clave(soup)  # panel "Datos clave" del nuevo layout
    cp = extraer_claves_puesto(soup)   # panel "Claves del puesto" (páginas por puesto)

    def cv(*etiquetas):
        for et in etiquetas:
            k = normalizar_clave(et)
            if k in clave and clave[k]:
                return clave[k]
        return ""

    titulo = texto_elemento(soup.find("h1"))

    institucion = cv("Institución", "Institucion") or extraer_campo(full, ["Institución", "Institucion"])
    if not institucion:
        institucion = extraer_breadcrumb(full)
    contrato = cv("Tipo de contrato") or extraer_campo(full, ["Tipo de contrato"])
    nivel = extraer_campo(full, ["Hay puestos para"])
    lugar = extraer_campo(full, ["Lugar de labores"]) or cp.get("lugardelabores", "")
    remuneracion = cv("Remuneración", "Remuneracion") or extraer_campo(full, ["Remuneración", "Remuneracion"]) or cp.get("remuneracion", "")
    carreras = extraer_campo(full, ["Hay plazas para"])
    plazas = (cv("Cantidad de plazas") or extraer_campo(full, ["Cantidad de plazas", "Número de vacantes", "Vacantes"])
              or cp.get("vacantedisponible", ""))

    # El h1 de las fichas va "Convocatoria SENASA ANALISTA EN SANIDAD..." o
    # "Convocatoria INEI OPERADOR... - UCAYALI": se quita el prefijo para dejar
    # el puesto limpio ("OPERADOR DE MANTENIMIENTO DE LOCAL DE EVALUACIÓN - UCAYALI").
    prefijo = re.match(r"^\s*[Cc]onvocatoria(?:\s+(?:N[°ºo]|N|Nro|numero)?\s*)?(?:\s+[A-ZÁÉÍÓÚÑ0-9]+)*?\s+", titulo)
    if prefijo and titulo.lower().find(institucion.lower()) <= len("convocatoria") + 12:
        resto = titulo[prefijo.end():].strip()
        if resto.lower().startswith(institucion.lower()):
            resto = resto[len(institucion):].lstrip("- —: ").strip()
        if resto:
            titulo = resto

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

    # Cronograma: "Del 14 de setiembre al 05 de octubre del 2026" -> inicio real
    # de postulación + fecha de fin de la convocatoria.
    inicio, fin_rango = fecha_rango_a_iso(full)

    # Vigencia: fecha límite ("esta vigente hasta el DD/MM/AAAA" o "Finaliza: dd de mmmm")
    vig = ""
    m = re.search(r"[Vv]igente hasta el\s*(\d{1,2})/(\d{1,2})/(\d{4})", full)
    if m:
        try:
            vig = date(int(m.group(3)), int(m.group(2)), int(m.group(1))).isoformat()
        except ValueError:
            vig = ""
    if not vig:
        vig = fecha_larga_a_iso(extraer_campo(full, ["Finaliza", "Fecha límite"]))
    if not vig:
        vig = fin_rango
    if not vig:
        vig = fecha_corta_a_iso(cp.get("plazoparapostular", ""))

    # Fecha de publicación ('25 sept 2026' en el nuevo layout, o forma larga)
    fecha = fecha_corta_a_iso(cv("Fecha de publicación"))
    if not fecha:
        fecha = fecha_larga_a_iso(extraer_campo(full, ["Fecha de Publicación", "Fecha de Publicacion"]))

    # Descripción (resumen amigable de la convocatoria)
    partes = []
    if contrato:
        partes.append(f"Contrato: {contrato}")
    if nivel:
        partes.append(f"Hay puestos para: {nivel}")
    if carreras:
        partes.append(f"Plazas para: {carreras}")
    if plazas:
        partes.append(f"Cantidad de plazas: {plazas}")
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

    enlaces = extraer_enlaces(soup)
    fila = {
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
        "inicio": inicio,
        "categoria": categoria,
        **BLANCOS,
    }
    # Enlaces directos de la convocatoria (mandan sobre el botón genérico:
    # la página los muestra como POSTULAR / VER BASES / CÓMO POSTULAR).
    fila["postular"] = enlaces["postular"]
    fila["ver_detalles"] = enlaces["bases"]
    fila["guia_registro"] = enlaces["video"]
    return fila


# --------------------------------------------------------------- GUARDADO

def guardar_json(datos):
    with open(ARCHIVO_JSON, "w", encoding="utf-8") as f:
        json.dump(datos, f, ensure_ascii=False, indent=2)
    print(f"[ok] JSON guardado: {ARCHIVO_JSON}")


def guardar_excel(datos):
    import openpyxl
    columnas = ["id", "titulo", "empresa", "ubicacion", "sueldo", "descripcion",
                "enlace", "whatsapp", "fecha", "fuente", "destacado", "visible",
                "ver_detalles", "funciones", "guia_registro", "postular", "vigencia",
                "inicio", "categoria"]
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Convocatorias"
    ws.append(columnas)
    for fila in datos:
        ws.append([fila.get(c, "") for c in columnas])
    anchos = [6, 45, 28, 22, 24, 70, 60, 8, 12, 26, 10, 8, 26, 26, 26, 42, 12, 12, 16]
    for i, w in enumerate(anchos, 1):
        ws.column_dimensions[chr(64 + i)].width = w
    wb.save(ARCHIVO_XLSX)
    print(f"[ok] Excel guardado: {ARCHIVO_XLSX} ({len(datos)} filas)")


# --------------------------------------------------------------- EJECUCIÓN

# Listados "trabajos por carrera" del sitio (pedido del usuario 2026-09-28):
# cada página lista ~30 PUESTOS vigentes (oportunidad-laboral-...html) con el
# mismo pill "Finaliza en N días". Número de páginas inicial leído de la paginación.
CARRERAS = [
    ("https://www.convocatoriasdetrabajo.com/carreras-profesionales-universitarias.php",
     "oportunidad-laboral-", 4, "Universitarios"),
    ("https://www.convocatoriasdetrabajo.com/carreras-profesionales-tecnicas.php",
     "oportunidad-laboral-", 2, "Técnicos"),
]


def listar_categoria(base, slug, paginas_max):
    """Url de los puestos VIGENTES de un listado por carrera (todas las páginas,
    hasta encontrar una sin tarjetas con pill vigente)."""
    urls = {}
    for p in range(1, paginas_max + 1):
        u = base if p == 1 else f"{base}?page={p}"
        html = descargar(u)
        if not html:
            continue
        nuevas = extraer_activas(html, u, slug)
        if not nuevas:
            break  # ya no hay vigentes en esta página
        urls.update(nuevas)
    return urls


def procesar(activas, resultados, categoria, max_detalles=0):
    """Scrapea el detalle de cada url y lo fusiona en `resultados`."""
    ordenadas = sorted(activas.items(),
                       key=lambda kv: (kv[1]["vigencia"], kv[0]), reverse=True)
    if max_detalles:
        ordenadas = ordenadas[:max_detalles]
    n = 0
    for pos, (url, info) in enumerate(ordenadas, 1):
        if url in resultados:
            continue
        print(f"[{categoria or 'Portada'} {pos}/{len(ordenadas)}] {url.split('/')[-1][:60]}")
        detalle = scrapear_detalle(url, info["region"], categoria)
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
        n += 1
        if n % 10 == 0:
            guardar_json(list(resultados.values()))
    return n


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

    resultados = {}
    procesar(activas, resultados, "", max_detalles)

    for base, slug, paginas, nombre in CARRERAS:
        print(f"--- Trabajos por carrera: {nombre} ---")
        urls = listar_categoria(base, slug, paginas)
        print(f"  puestos vigentes en {nombre}: {len(urls)}")
        procesar(urls, resultados, nombre, max_detalles or 0)

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