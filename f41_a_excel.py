#!/usr/bin/env python3
"""Extrae la tabla de items de un Pedido de Cotizacion F.41
(Gobierno de la Provincia de Formosa) y la vuelca a un Excel.

Uso:
    python3 f41_a_excel.py archivo1.pdf archivo2.pdf ...
    python3 f41_a_excel.py carpeta_con_pdfs/

Por cada PDF de entrada se genera un .xlsx con el mismo nombre, en la
misma carpeta que el PDF.
"""

import sys
from pathlib import Path

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter

# Coordenadas (en puntos) de las columnas de la tabla "DETALLE DE ITEMS"
# del formulario F.41. Son fijas porque el formulario es una plantilla
# oficial con posiciones constantes.
COL_RG_MAX = 95
COL_CODIGO_MAX = 168
COL_DESC_MAX = 384
COL_CANT_MAX = 420
LEFT_BORDER_X = 70.9
LEFT_BORDER_TOL = 2

# Sólo se conservan los items cuyo código empieza con alguno de estos prefijos.
PREFIJOS_CODIGO = ("4.01.008", "4.01.018")


def bandas_de_filas(page):
    """Detecta las bandas verticales (top, bottom) de cada fila de la
    tabla a partir de los segmentos del borde izquierdo."""
    segmentos = [
        l
        for l in page.lines
        if abs(l["x0"] - l["x1"]) < 0.5 and abs(l["x0"] - LEFT_BORDER_X) < LEFT_BORDER_TOL
    ]
    segmentos.sort(key=lambda l: l["top"])
    return [(s["top"], s["bottom"]) for s in segmentos]


def extraer_filas_pagina(page):
    palabras = page.extract_words()
    filas = []
    for top, bottom in bandas_de_filas(page):
        en_banda = [w for w in palabras if top - 0.5 <= w["top"] < bottom - 0.5]
        if not en_banda:
            continue

        rg_palabras = [w for w in en_banda if w["x0"] < COL_RG_MAX]
        codigo_palabras = [w for w in en_banda if COL_RG_MAX <= w["x0"] < COL_CODIGO_MAX]
        # La columna "Cantidad" está alineada a la derecha, así que un
        # número largo (ej. "152064") puede empezar (x0) antes del límite
        # de la columna de Descripción aunque termine (x1) bien adentro de
        # la columna de Cantidad. Por eso esta columna se distingue por su
        # borde derecho (x1) y no por el izquierdo (x0).
        desc_palabras = [
            w
            for w in en_banda
            if COL_CODIGO_MAX <= w["x0"] < COL_DESC_MAX and w["x1"] <= COL_DESC_MAX
        ]
        cant_palabras = [
            w
            for w in en_banda
            if w["x0"] >= COL_CODIGO_MAX and COL_DESC_MAX < w["x1"] <= COL_CANT_MAX
        ]

        rg_texto = "".join(w["text"] for w in rg_palabras)
        if not rg_texto.isdigit():
            # No es una fila de item real (encabezado de columnas, fila
            # de "Total", etc.)
            continue

        codigo = "".join(w["text"] for w in codigo_palabras)
        desc_palabras.sort(key=lambda w: (round(w["top"], 1), w["x0"]))
        descripcion = " ".join(w["text"] for w in desc_palabras)
        cant_texto = "".join(w["text"] for w in cant_palabras)
        try:
            cantidad = int(cant_texto)
        except ValueError:
            cantidad = cant_texto

        filas.append(
            {
                "rg": int(rg_texto),
                "codigo": codigo,
                "descripcion": descripcion,
                "cantidad": cantidad,
            }
        )
    return filas


def extraer_pdf(pdf_path: Path):
    with pdfplumber.open(pdf_path) as pdf:
        filas = []
        for page in pdf.pages:
            filas.extend(extraer_filas_pagina(page))
    return [f for f in filas if f["codigo"].startswith(PREFIJOS_CODIGO)]


def escribir_excel(filas: list, salida: Path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Cotizacion"

    negrita = Font(bold=True)
    borde = Border(*(Side(style="thin"),) * 4)

    fila = 1
    encabezados_tabla = ["Nº", "Código", "Descripción", "Cantidad", "Precio Unitario", "Total"]
    fila_encabezado_tabla = fila
    for col, titulo in enumerate(encabezados_tabla, start=1):
        c = ws.cell(row=fila, column=col, value=titulo)
        c.font = negrita
        c.border = borde
        c.alignment = Alignment(horizontal="center")
    fila += 1

    primera_fila_datos = fila
    for item in filas:
        ws.cell(row=fila, column=1, value=item["rg"]).border = borde
        ws.cell(row=fila, column=2, value=item["codigo"]).border = borde
        c_desc = ws.cell(row=fila, column=3, value=item["descripcion"])
        c_desc.border = borde
        c_desc.alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(row=fila, column=4, value=item["cantidad"]).border = borde
        c_pu = ws.cell(row=fila, column=5)
        c_pu.border = borde
        c_pu.number_format = '"$"#,##0.00'
        c_total = ws.cell(row=fila, column=6, value=f"=D{fila}*E{fila}")
        c_total.border = borde
        c_total.number_format = '"$"#,##0.00'
        fila += 1
    ultima_fila_datos = fila - 1

    ws.cell(row=fila, column=5, value="Total:").font = negrita
    c_total_general = ws.cell(row=fila, column=6, value=f"=SUM(F{primera_fila_datos}:F{ultima_fila_datos})")
    c_total_general.font = negrita
    c_total_general.number_format = '"$"#,##0.00'

    anchos = {"A": 6, "B": 16, "C": 60, "D": 10, "E": 16, "F": 16}
    for letra, ancho in anchos.items():
        ws.column_dimensions[letra].width = ancho

    ws.freeze_panes = f"A{fila_encabezado_tabla + 1}"

    wb.save(salida)


def procesar(pdf_path: Path):
    print(f"Procesando: {pdf_path.name}")
    filas = extraer_pdf(pdf_path)
    if not filas:
        print(f"  ADVERTENCIA: no se detectaron items en {pdf_path.name}")
    salida = pdf_path.with_suffix(".xlsx")
    escribir_excel(filas, salida)
    print(f"  -> {salida} ({len(filas)} items)")


def main(argv):
    if not argv:
        print(__doc__)
        return 1

    pdfs = []
    for arg in argv:
        p = Path(arg)
        if p.is_dir():
            pdfs.extend(sorted(p.glob("*.pdf")))
        elif p.is_file():
            pdfs.append(p)
        else:
            print(f"AVISO: no existe el archivo/carpeta: {arg}")

    if not pdfs:
        print("No se encontraron archivos PDF para procesar.")
        return 1

    errores = 0
    for pdf_path in pdfs:
        try:
            procesar(pdf_path)
        except Exception as exc:  # noqa: BLE001 - queremos seguir con el resto
            errores += 1
            print(f"  ERROR procesando {pdf_path.name}: {exc}")

    return 1 if errores else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
