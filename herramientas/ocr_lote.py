# -*- coding: utf-8 -*-
"""OCR en LOTE de las imagenes nuevas de `images/` (tipo IMG-AAAAMMDD-WA00xx.jpg).

Mejora sobre `ocr_multi.py`: amplía la imagen (los carteles de la agencia son
chiquitos y el texto sale ilegible), la parte en 2-3 recuadros, aplica umbral
para tapar el fondo de colores y corre tesseract en psm 6 y 11 (densa y en
bloques). Los resultados quedan en `herramientas/ocr_lote.json`.

Uso:
    python3 herramientas/ocr_lote.py IMG-20261021          # todas las de ese lote
    python3 herramientas/ocr_lote.py IMG-20261021-WA0066   # solo una
    python3 herramientas/ocr_lote.py --fuente /ruta/otra.jpg
"""
import json
import os
import re
import subprocess
import sys
from concurrent.futures import ThreadPoolExecutor

from PIL import Image, ImageOps

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
IMGS = os.path.join(BASE, "images")
TMP = os.path.join(BASE, "herramientas", "_ocr_tmp")
SALIDA = os.path.join(BASE, "herramientas", "ocr_lote.json")
LARGO = 2200.0          # lado mayor al que se amplía
UMBRAL = 135            # 0-255: > umbral = tinta (negro), < = fondo (blanco)
REC = [(0.00, 0.00, 1.00, 0.55), (0.00, 0.45, 1.00, 1.00)]   # recuadros (x1,y1,x2,y2)
PSM = ["6", "11"]


def recortes(ruta):
    """Amplía la imagen y la parte en recuadros; devuelve la lista de PNG."""
    os.makedirs(TMP, exist_ok=True)
    im = Image.open(ruta).convert("RGB")
    w, h = im.size
    esc = LARGO / max(w, h)
    im = im.resize((int(w * esc), int(h * esc)), Image.LANCZOS)
    W, H = im.size
    rutas = []
    base = os.path.splitext(os.path.basename(ruta))[0]
    for i, (x1, y1, x2, y2) in enumerate(REC):
        c = im.crop((int(x1 * W), int(y1 * H), int(x2 * W), int(y2 * H)))
        g = ImageOps.autocontrast(ImageOps.grayscale(c), cutoff=1)
        p = os.path.join(TMP, f"{base}__r{i}.png")
        g.point(lambda v: 255 if v > UMBRAL else 0).convert("RGB").save(p, dpi=(300, 300))
        rutas.append(p)
    return rutas


def ocr(png, psm):
    try:
        r = subprocess.run(["tesseract", png, "stdout", "-l", "spa", "--psm", psm],
                           capture_output=True, text=True, timeout=280)
        return r.stdout.strip()
    except subprocess.TimeoutExpired:
        return "[timeout]"


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    if not args:
        print(__doc__)
        return
    if args[0] == "--fuente":
        archivos = [args[1]]
    else:
        archivos = sorted(f for f in os.listdir(IMGS)
                          if f.startswith(args[0]) and f.lower().endswith((".jpg", ".jpeg", ".png")))
    print("imagenes:", len(archivos))
    datos = {}
    for f in archivos:
        ruta = os.path.join(IMGS, f) if not f.startswith("/") else f
        if not os.path.exists(ruta):
            print("no existe:", ruta)
            continue
        salida = {}
        for png in recortes(ruta):
            for psm in PSM:
                salida[os.path.basename(png) + ":" + psm] = ocr(png, psm)
        datos[os.path.basename(ruta)] = salida
        print("ok", os.path.basename(ruta))
    with open(SALIDA, "w", encoding="utf-8") as fh:
        json.dump(datos, fh, ensure_ascii=False, indent=1)
    print("Guardado", SALIDA)


if __name__ == "__main__":
    main()
