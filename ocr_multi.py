#!/usr/bin/env python3
"""OCR multi-modo para anuncios (FOTOS de images/).

Algunos carteles llevan un SELLO/LOGOTIPO de empresa (p.ej. "EMPRESA APU LLALLAWA")
que el OCR estándar pierde -> se lee distinto según el modo (psm). Este script ejecuta
tesseract en varios modos y compara los resultados para extraer la empresa real.

Uso:
  python3 ocr_multi.py "images/WhatsApp Image 2026-09-15 at 18.46.16.jpeg"
  python3 ocr_multi.py --recorrer images         # todas las imágenes de la carpeta
"""
import os
import re
import subprocess
import sys

MODOS = ["3", "6", "11", "4"]
# Señales de un sello/logotipo de empresa (el nombre real de la empresa).
CLAVES_EMPRESA = ["EMPRESA", "S.A.C", "SAC", "CONTRATISTA", "CONSTRUCTORA",
                  "MINERA", "CORPORACI", "INVERSIONES", "TRANSPORTES", "E.I.R.L"]


def ocr(path, psm):
    try:
        r = subprocess.run(["tesseract", path, "stdout", "-l", "spa", "--psm", psm],
                           capture_output=True, text=True, timeout=120)
        return r.stdout or ""
    except Exception as e:
        return f"[error] {e}"


def detectar_sellos(textos):
    """Busca las CLaves de empresa en cada resultado y reporta el fragmento."""
    hallazgos = []
    for psm, txt in textos.items():
        for clave in CLAVES_EMPRESA:
            patron = re.compile(r"[A-ZÁÉÍÓÚÑ0-9][^\n.]{0,30}?" + clave +
                                r"[ \t]*[A-ZÁÉÍÓÚÑ0-9][^\n.]{0,40}", re.IGNORECASE)
            for m in patron.finditer(txt):
                seg = re.sub(r"\s+", " ", m.group(0)).strip()
                if seg and seg not in hallazgos:
                    hallazgos.append(seg)
    return hallazgos


def analizar(path):
    print("=" * 70)
    print("IMAGEN:", path)
    print("=" * 70)
    textos = {m: ocr(path, m) for m in MODOS}
    unificados = " ".join(textos.values())
    for psm, txt in textos.items():
        limpio = " | ".join(l.strip() for l in txt.splitlines() if l.strip()[:25])
        print(f"\n-- psm {psm} --")
        print(txt.strip()[:900] if txt.strip() else "(vacío)")
    sellos = detectar_sellos(textos)
    if sellos:
        print("\n>>> POSIBLES SELLOS/EMPRESAS detectados:")
        for s in sellos:
            print("    -", s)
    # números de contacto
    nums = sorted(set(re.findall(r"(?<![0-9])9\d{8}(?![0-9])", unificados)))
    if nums:
        print("\n>>> NÚMEROS (9 dígitos):", nums)
    print()


def recorrer(carpeta):
    archivos = sorted([os.path.join(carpeta, f) for f in os.listdir(carpeta)
                       if f.lower().endswith((".jpeg", ".jpg", ".png"))])
    for a in archivos:
        analizar(a)


if __name__ == "__main__":
    args = sys.argv[1:]
    if "--recorrer" in args:
        carpeta = args[args.index("--recorrer") + 1] if len(args) > args.index("--recorrer") + 1 else "images"
        recorrer(carpeta)
    elif args:
        for p in args:
            analizar(p)
    else:
        print(__doc__)