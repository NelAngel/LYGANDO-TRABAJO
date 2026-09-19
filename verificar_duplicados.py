#!/usr/bin/env python3
"""Anti-duplicados de LYGANDO TRABAJOS (regla del usuario 2026-09-17).

REGLA:
  - Si TODOS los datos se repiten (PUESTO + LUGAR + NÚMERO) -> IGNORAR: es la
    misma oferta; no se registra de nuevo (la página tampoco la repite).
  - Si algo cambió (el PUESTO, el LUGAR o el NÚMERO) -> AVISAR al usuario para
    que decida (puede ser otra vacante real del mismo negocio).
  Escanea AMBAS fuentes (ofertas.xlsx de fotos + convocatorias.xlsx del scraping)
  para detectar también duplicados ENTRE fuentes. Los números se normalizan
  (se quita +51 / 0 inicial) para no tratar la misma oferta como distinta.

Uso:
  python3 verificar_duplicados.py                         -> revisa imágenes repetidas
                                                              y filas duplicadas/parecidas
  python3 verificar_duplicados.py --filtrar               -> marca visible=no las duplicadas
                                                              exactas (conserva la reciente)
  python3 verificar_duplicados.py --ver "PUESTO" "LUGAR" "NÚMERO"
                                                          -> compara una oferta nueva
                                                             (recién leída con OCR) contra
                                                             lo ya publicado

DETECCIÓN DE IMÁGENES (siempre, además del SHA-1):
  - SHA-1 igual       -> MISMA foto byte a byte (renombrada) -> IGNORAR.
  - Píxeles 64x64 <=6% -> MISMA foto re-enviada/re-comprimida (cambió el archivo pero
                          es el mismo cartel) -> IGNORAR.
  - Píxeles 64x64 6-10% -> PARECIDA (plantilla parecida: revisar antes de registrar).
  Se usa comparación de píxeles (no pHash: las plantillas de la agencia confunden todo).
  Requiere Pillow:  pip install pillow
"""
import glob
import hashlib
import os
import sys
import unicodedata
from collections import defaultdict

import openpyxl

try:
    from PIL import Image
    _PILLOW = True
except ImportError:
    Image = None
    _PILLOW = False

OFERTAS = "ofertas.xlsx"
CONVOCATORIAS = "convocatorias.xlsx"
CARPETA_IMAGES = "images"


def normalizar(s):
    t = str(s or "").lower()
    t = unicodedata.normalize("NFD", t)
    return "".join(c for c in t if unicodedata.category(c) != "Mn").strip()


def numero_de(v):
    """Dígitos del número, normalizados: se quita el código de país 51 y un 0
    inicial -> un mismo teléfono escrito '997...', '+51 997...' o '09997...' es igual."""
    d = "".join(c for c in str(v or "") if c.isdigit())
    for _ in range(2):
        if d.startswith("0051"):
            d = d[4:]
        if d.startswith("51") and len(d) >= 11:
            d = d[2:]
        if d.startswith("0") and len(d) >= 10:
            d = d[1:]
    return d


def sha1_de(archivo):
    h = hashlib.sha1()
    with open(archivo, "rb") as f:
        for bloque in iter(lambda: f.read(65536), b""):
            h.update(bloque)
    return h.hexdigest()


def imagenes_repetidas():
    archivos = sorted(glob.glob(os.path.join(CARPETA_IMAGES, "*")))
    por_hash = defaultdict(list)
    for a in archivos:
        if os.path.isfile(a):
            por_hash[sha1_de(a)].append(a)
    return {h: lista for h, lista in por_hash.items() if len(lista) > 1}


# Umbrales de similitud por píxeles (calibrados con los carteles reales 2026-09-17:
# el duplicado conocido da 0.0% y los carteles distintos de la misma plantilla >=10.4%).
DIF_MISMA_FOTO = 6.0     # menor o igual -> IGNORAR (misma foto re-enviada)
DIF_PARECIDA = 10.0      # menor que esto y mayor que DIF_MISMA_FOTO -> REVISAR


def _mini_64(ruta):
    """Miniatura 64x64 en escala de grises (cache para comparaciones rápidas)."""
    img = Image.open(ruta).convert("L")
    return list(img.resize((64, 64), Image.LANCZOS).getdata())


def diferencia_pixeles(a, b, tol=24):
    """% de píxeles que difieren >tol (0-100). 0 = idénticas."""
    n = len(a)
    dif = sum(1 for i in range(n) if abs(a[i] - b[i]) > tol)
    return 100.0 * dif / n


def imagenes_similares():
    """Devuelve (mismas_foto, parecidas).

    mismas_foto: pares (d, a, b) con dif <= DIF_MISMA_FOTO -> IGNORAR.
    parecidas  : pares (d, a, b) con DIF_MISMA_FOTO < dif < DIF_PARECIDA -> REVISAR.
    Se saltan los pares ya detectados por SHA-1 (byte a byte).
    """
    archivos = sorted(glob.glob(os.path.join(CARPETA_IMAGES, "*")))
    archivos = sorted(a for a in archivos if os.path.isfile(a))
    if not _PILLOW or len(archivos) < 2:
        return [], []

    por_hash = defaultdict(list)
    for a in archivos:
        por_hash[sha1_de(a)].append(a)
    hash_duplicado = {h for h, lista in por_hash.items() if len(lista) > 1}
    hash_de = {a: sha1_de(a) for a in archivos}

    mins = {a: _mini_64(a) for a in archivos}
    mismas_foto = []
    parecidas = []
    for i in range(len(archivos)):
        for j in range(i + 1, len(archivos)):
            a, b = archivos[i], archivos[j]
            if hash_de[a] in hash_duplicado and hash_de[b] == hash_de[a]:
                continue
            d = diferencia_pixeles(mins[a], mins[b])
            if d <= DIF_MISMA_FOTO:
                mismas_foto.append((d, a, b))
            elif d < DIF_PARECIDA:
                parecidas.append((d, a, b))
    mismas_foto.sort()
    parecidas.sort()
    return mismas_foto, parecidas


def filas_de(archivo, origen):
    """Lee las filas de un Excel (convocatorias.xlsx también, para detectar
    duplicados ENTRE fuentes: foto de la agencia vs convocatoria ONPE)."""
    wb = openpyxl.load_workbook(archivo)
    ws = wb.active
    cab = [c.value for c in ws[1]]
    idx = {n: i for i, n in enumerate(cab)}
    filas = []
    for fila in ws.iter_rows(min_row=2):
        vals = list(fila)
        if vals[0].value is None:
            continue
        filas.append({
            "archivo": archivo,
            "origen": origen,
            "indice": fila[0].row,
            "id": vals[idx["id"]].value,
            "titulo": vals[idx["titulo"]].value,
            "empresa": vals[idx["empresa"]].value if "empresa" in idx else "",
            "ubicacion": vals[idx["ubicacion"]].value if "ubicacion" in idx else "",
            "whatsapp": vals[idx["whatsapp"]].value if "whatsapp" in idx else "",
            "descripcion": vals[idx["descripcion"]].value if "descripcion" in idx else "",
            "visible": vals[idx["visible"]].value if "visible" in idx else "si",
            "fecha": vals[idx["fecha"]].value if "fecha" in idx else ""})
    return wb, filas


def identificar(r):
    return f"id {r['id']} ({r['origen']})"


def es_visible(r):
    return not (not r["visible"] or str(r["visible"]).strip().lower() in ("no", "0", "false"))


def clave_completa(r):
    """Clave de 'oferta repetida COMPLETA': puesto + número + lugar + empresa."""
    tit = normalizar(r["titulo"])
    if not tit:
        return ""
    return "|".join([tit, numero_de(r["whatsapp"]), normalizar(r["ubicacion"]),
                     normalizar(r["empresa"])])


def visibles(filas):
    return [r for r in filas if es_visible(r)]


def analizar_filas(filas):
    """Devuelve (duplicados_exactos, avisos).

    duplicados_exactos: filas con TODOS los datos iguales -> IGNORAR.
    avisos: filas que coinciden en algo (número o puesto+lugar) pero con
            algún cambio (puesto, lugar o número) -> AVISAR.
    """
    vis = visibles(filas)

    exactos = defaultdict(list)
    for r in vis:
        c = clave_completa(r)
        if c:
            exactos[c].append(r)
    exactos = {c: g for c, g in exactos.items() if len(g) > 1}

    avisos = []

    # Mismo NÚMERO pero cambió el PUESTO o el LUGAR.
    numero_grupos = defaultdict(list)
    for r in vis:
        n = numero_de(r["whatsapp"])
        if n:
            numero_grupos[n].append(r)
    for n, g in numero_grupos.items():
        if len(g) < 2:
            continue
        tits = {normalizar(r["titulo"]) for r in g}
        lugs = {normalizar(r["ubicacion"]) for r in g}
        if len(tits) > 1 or len(lugs) > 1:
            ids = ", ".join(identificar(r) for r in g)
            cambios = []
            if len(tits) > 1:
                cambios.append("el PUESTO")
            if len(lugs) > 1:
                cambios.append("el LUGAR")
            avisos.append(f"  AVISO: mismo número ({n}) en {ids} pero cambió "
                          f"{' y '.join(cambios)} -> no es un duplicado exacto, revisar.")

    # Mismo PUESTO + LUGAR pero cambió el NÚMERO.
    titlug_grupos = defaultdict(list)
    for r in vis:
        tl = (normalizar(r["titulo"]), normalizar(r["ubicacion"]))
        if tl[0] and tl[1]:
            titlug_grupos[tl].append(r)
    for tl, g in titlug_grupos.items():
        nums = {numero_de(r["whatsapp"]) for r in g}
        if len(g) > 1 and len(nums) > 1:
            ids = ", ".join(identificar(r) for r in g)
            avisos.append(f"  AVISO: mismo puesto y lugar en {ids} "
                          f"({g[0]['titulo']} / {g[0]['ubicacion']}) pero cambió el NÚMERO "
                          f"-> pueden ser dos vacantes distintas, revisar.")

    return exactos, avisos


def verificar_candidata(filas):
    """Compara una oferta nueva (recién leída con OCR) contra lo publicado."""
    args = [a for a in sys.argv[2:] if not a.startswith("--")]
    if len(args) < 3:
        print("  Uso: python3 verificar_duplicados.py --ver \"PUESTO\" \"LUGAR\" \"NUMERO\"")
        return
    titulo, lugar, numero = args[0], args[1], args[2]
    num = numero_de(numero)
    tit, lug = normalizar(titulo), normalizar(lugar)

    exacto = None
    avisos = []
    for r in visibles(filas):
        rt = normalizar(r["titulo"])
        rl = normalizar(r["ubicacion"])
        rn = numero_de(r["whatsapp"])

        todo_igual = (tit and rt == tit) and (lug and rl == lug) and (not num or rn == num)
        if todo_igual:
            exacto = r
            continue

        # Mismo número pero ha cambiado el puesto o el lugar.
        if num and rn and num == rn and (rt != tit or rl != lug):
            cambios = []
            if rt != tit:
                cambios.append("el PUESTO")
            if rl != lug:
                cambios.append("el LUGAR")
            avisos.append(f"  AVISO: mismo número ({num}) que {identificar(r)} pero cambió "
                          f"{' y '.join(cambios)} -> ¿es la misma oferta con datos nuevos?")

        # Mismo puesto y lugar pero cambió el número.
        elif tit and rt == tit and lug and rl == lug:
            if num and rn and rn != num:
                avisos.append(f"  AVISO: mismo puesto y lugar que {identificar(r)} pero cambió "
                              f"el NÚMERO: {rn} -> {num}")
            elif not num:
                avisos.append(f"  AVISO: mismo puesto y lugar que {identificar(r)} pero la nueva "
                              f"foto no trae número.")

    print("\n== Verificando la oferta nueva ==")
    print(f"  PUESTO : {titulo}")
    print(f"  LUGAR  : {lugar}")
    print(f"  NÚMERO : {numero}")
    if exacto:
        print(f"  -> IGNORAR: ya está publicada ({identificar(exacto)}, SUBIDA "
              f"el {exacto['fecha']}). NO registrarla de nuevo.")
    for a in avisos:
        print(a)
    if not exacto and not avisos:
        print("  -> OK: sin conflictos, se puede registrar.")
    elif avisos:
        print("  -> Revisar los avisos: si es la misma foto con datos corregidos, borra la "
              "anterior; si es otra vacante, registra normalmente.")


def main():
    quiere_filtrar = "--filtrar" in sys.argv
    quiere_ver = "--ver" in sys.argv

    print(f"== Imágenes repetidas en {CARPETA_IMAGES}/ ==")
    rep = imagenes_repetidas()
    for h, lista in rep.items():
        print(f"  IGNORAR ({h[:8]}), es LA MISMA foto (registrar UNA sola vez):")
        for a in lista:
            print(f"    - {os.path.basename(a)}")
    if not rep:
        print("  Sin imágenes repetidas (todas distintas).")

    if _PILLOW:
        mismas, parecidas = imagenes_similares()
        if mismas:
            print("\n  MISMA FOTO re-enviada/re-comprimida (<=%.0f%% de diferencia)"
                  % DIF_MISMA_FOTO)
            print("  -> IGNORAR: es el mismo cartel aunque el archivo cambió:")
            for d, a, b in mismas:
                print(f"    - {d:5.1f}%  {os.path.basename(a)}  <->  {os.path.basename(b)}")
        if parecidas:
            print(f"\n  PARECIDAS (entre {DIF_MISMA_FOTO:.0f}% y {DIF_PARECIDA:.0f}% de "
                  f"diferencia) -> REVISAR antes de registrar:")
            for d, a, b in parecidas:
                print(f"    - {d:5.1f}%  {os.path.basename(a)}  <->  {os.path.basename(b)}")
        if not mismas and not parecidas:
            print("  Sin imágenes similares (solo se marca el duplicado real por SHA-1).")
    else:
        print("\n  (aviso: no hay Pillow, no se compara similitud por píxeles ->")
        print("   pip install pillow)")

    print(f"\n== Filas ({OFERTAS} + {CONVOCATORIAS}) ==")
    wb_fotos, filas_fotos = filas_de(OFERTAS, "Foto")
    _, filas_onpe = filas_de(CONVOCATORIAS, "PortalTrabajo")
    filas = filas_fotos + filas_onpe
    exactos, avisos = analizar_filas(filas)

    if exactos:
        print("  DUPLICADOS EXACTOS (mismo puesto + lugar + número) -> IGNORAR:")
        for clave, grupo in exactos.items():
            ids = ", ".join(identificar(r) for r in grupo)
            print(f"    - {ids} | {grupo[0]['titulo']} / {grupo[0]['ubicacion']} / "
                  f"{grupo[0]['whatsapp']}")
    else:
        print("  Sin duplicados exactos.")

    if avisos:
        print("  AVISOS (algo cambió entre ofertas parecidas):")
        for a in avisos:
            print(a)
    else:
        print("  Sin ofertas parecidas con datos cambiados.")

    if quiere_filtrar and exactos:
        ws = wb_fotos.active
        col_visible = None
        for nombre, i in {c.value: n for n, c in enumerate(ws[1])}.items():
            if nombre == "visible":
                col_visible = i + 1
        borrados = 0
        for clave, grupo in exactos.items():
            grupo_ordenado = sorted(grupo, key=lambda x: str(x["fecha"] or ""))
            for r in grupo_ordenado[:-1]:
                if r["archivo"] != OFERTAS or not col_visible:
                    continue
                ws.cell(row=r["indice"], column=col_visible).value = "no"
                borrados += 1
        ws.parent.save(OFERTAS)
        print(f"\n  Marcadas visible=no en {OFERTAS}: {borrados} fila(s) "
              f"(conservadas las recientes).")

    if quiere_ver:
        verificar_candidata(filas)


if __name__ == "__main__":
    main()