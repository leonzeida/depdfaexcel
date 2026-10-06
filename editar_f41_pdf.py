#!/usr/bin/env python3
"""Genera una versión EDITADA del PDF F.41 original: mantiene el
membrete, el texto legal y la firma tal cual están en el PDF que se
subió, y reemplaza únicamente la tabla de ítems por la que el usuario
armó en la grilla web (agregando, sacando, reordenando y editando
renglones, incluidos precio unitario y total).

A diferencia de un documento generado de cero: el membrete y el texto
legal de la página 1 (y el bloque corto que se repite en cada página) se
recortan del PDF original como IMÁGENES y se reusan tal cual — no se
reconstruyen a mano — para garantizar que se vean exactamente iguales.
Solo la tabla de ítems y el total final se dibujan de nuevo, con los
datos editados.

Se mantiene separado de f41_a_excel.py (que es de extracción/Excel) y es
el reemplazo de generar_pedido_pdf.py de la ronda anterior (esa versión
generaba un documento propio de Zeid Medical, no el F.41 editado).
"""

import io
import re
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import pdfplumber
from num2words import num2words
from reportlab.lib.utils import ImageReader
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    NextPageTemplate,
    PageTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib import colors

from f41_a_excel import COL_CANT_MAX, COL_CODIGO_MAX, COL_DESC_MAX, COL_RG_MAX, LEFT_BORDER_X, bandas_de_filas

ZONA_HORARIA_ARGENTINA = ZoneInfo("America/Argentina/Buenos_Aires")

# Separación observada, en el PDF real, entre el final del bloque corto
# de encabezado (membrete + 3 líneas de expediente) y el inicio de la
# primera fila de ítems en una página de continuación.
PADDING_DEBAJO_MINI_ENCABEZADO = 14

MARGEN_INFERIOR = 30  # puntos, deja lugar al pie de página


def _formatear_moneda(valor) -> str:
    """1234567.89 -> "1.234.567,89" (separador de miles con punto, coma
    decimal, como en toda la planilla de Excel del comparador)."""
    if valor is None or valor == "":
        return ""
    texto = f"{float(valor):,.2f}"
    return texto.replace(",", "_").replace(".", ",").replace("_", ".")


def _monto_en_palabras(valor: float) -> str:
    centavos_totales = round(float(valor) * 100)
    entero, centavos = divmod(centavos_totales, 100)
    return f"SON PESOS: {num2words(entero, lang='es').upper()} PESOS CON {centavos:02d}/100.-"


def _analizar_pdf_original(pdf_path: Path) -> dict:
    """Lee el PDF que subió el usuario y saca todo lo necesario para
    poder reconstruirlo: tamaño de página, dónde termina el encabezado
    (completo en la página 1, corto en las páginas siguientes), el
    recorte de esas dos zonas como imagen, el ancho real de la tabla y
    el texto del pie de página.
    """
    with pdfplumber.open(pdf_path) as pdf:
        p1 = pdf.pages[0]
        ancho, alto = p1.width, p1.height

        lineas_h = sorted(
            {round(l["top"], 1) for l in p1.lines if abs(l["top"] - l["bottom"]) < 0.5}
        )
        fin_mini_encabezado = lineas_h[1] if len(lineas_h) > 1 else 140.0

        bandas_p1 = bandas_de_filas(p1)
        fin_encabezado_pagina1 = bandas_p1[0][0] if bandas_p1 else fin_mini_encabezado + PADDING_DEBAJO_MINI_ENCABEZADO

        borde_derecho = max(
            (l["x1"] for l in p1.lines if abs(l["top"] - l["bottom"]) < 0.5),
            default=ancho - LEFT_BORDER_X,
        )

        def _recortar(hasta_y: float) -> bytes:
            imagen = p1.crop((0, 0, ancho, hasta_y)).to_image(resolution=200)
            buffer = io.BytesIO()
            imagen.save(buffer, format="PNG")
            return buffer.getvalue()

        imagen_pagina1 = _recortar(fin_encabezado_pagina1)
        imagen_mini = _recortar(fin_mini_encabezado)

        texto_p1 = p1.extract_text() or ""
        coincidencia = re.search(r"(Usuario .+? P[aá]gina)\s*\d+", texto_p1)
        pie_pagina = coincidencia.group(1) if coincidencia else "Página"

    return {
        "ancho": ancho,
        "alto": alto,
        "fin_encabezado_pagina1": fin_encabezado_pagina1,
        "fin_mini_encabezado": fin_mini_encabezado,
        "imagen_pagina1": imagen_pagina1,
        "imagen_mini": imagen_mini,
        "borde_derecho": borde_derecho,
        "pie_pagina": pie_pagina,
    }


def _anchos_columnas(borde_derecho: float) -> list:
    """Reusa los mismos límites de columna que ya usa la extracción
    (f41_a_excel.py), para que la tabla nueva quede alineada en la misma
    posición horizontal que la original. Precio unitario/Total se
    reparten el resto del ancho de la tabla por la mitad."""
    resto = borde_derecho - COL_CANT_MAX
    mitad = resto / 2
    return [
        COL_RG_MAX - LEFT_BORDER_X,
        COL_CODIGO_MAX - COL_RG_MAX,
        COL_DESC_MAX - COL_CODIGO_MAX,
        COL_CANT_MAX - COL_DESC_MAX,
        mitad,
        mitad,
    ]


def _dibujar_encabezado(imagen_bytes: bytes, alto_imagen: float, ancho_pagina: float, alto_pagina: float, pie_pagina: str):
    lector = ImageReader(io.BytesIO(imagen_bytes))

    def _onpage(canvas, doc):
        canvas.saveState()
        canvas.drawImage(
            lector, 0, alto_pagina - alto_imagen,
            width=ancho_pagina, height=alto_imagen,
            preserveAspectRatio=False, mask="auto",
        )
        canvas.setFont("Helvetica", 7)
        canvas.drawString(LEFT_BORDER_X, 20, f"{pie_pagina} {doc.page}")
        canvas.restoreState()

    return _onpage


def generar_f41_editado(pdf_original_path: Path, filas: list, salida: Path):
    """`filas`: lista de dicts con "rg", "codigo", "descripcion",
    "cantidad", "precio_unitario", "total" (mismo criterio que la ronda
    anterior: los valores de precio/total ya vienen resueltos del
    frontend, no se recalculan acá).
    """
    info = _analizar_pdf_original(pdf_original_path)
    ancho, alto = info["ancho"], info["alto"]
    borde_derecho = info["borde_derecho"]

    doc = BaseDocTemplate(
        str(salida),
        pagesize=(ancho, alto),
        leftMargin=0, rightMargin=0, topMargin=0, bottomMargin=0,
        title="F.41 editado",
    )

    alto_disponible_p1 = (alto - info["fin_encabezado_pagina1"]) - MARGEN_INFERIOR
    alto_disponible_cont = (alto - info["fin_mini_encabezado"] - PADDING_DEBAJO_MINI_ENCABEZADO) - MARGEN_INFERIOR
    ancho_frame = borde_derecho - LEFT_BORDER_X

    frame_pagina1 = Frame(
        LEFT_BORDER_X, MARGEN_INFERIOR, ancho_frame, alto_disponible_p1,
        id="f1", leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0,
    )
    frame_continuacion = Frame(
        LEFT_BORDER_X, MARGEN_INFERIOR, ancho_frame, alto_disponible_cont,
        id="fc", leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0,
    )

    doc.addPageTemplates([
        PageTemplate(
            id="pagina1", frames=[frame_pagina1],
            onPage=_dibujar_encabezado(info["imagen_pagina1"], info["fin_encabezado_pagina1"], ancho, alto, info["pie_pagina"]),
        ),
        PageTemplate(
            id="continuacion", frames=[frame_continuacion],
            onPage=_dibujar_encabezado(info["imagen_mini"], info["fin_mini_encabezado"], ancho, alto, info["pie_pagina"]),
        ),
    ])

    estilos = getSampleStyleSheet()
    estilo_celda = ParagraphStyle("celda", parent=estilos["Normal"], fontName="Helvetica", fontSize=8, leading=9.5)
    estilo_celda_bold = ParagraphStyle("celda_bold", parent=estilo_celda, fontName="Helvetica-Bold")

    encabezados = ["Rg", "Código", "Descripción", "Cant.", "P.Unit.", "Total"]
    filas_tabla = [[Paragraph(h, estilo_celda_bold) for h in encabezados]]
    total_general = 0.0
    for fila in filas:
        total = fila.get("total")
        if total not in (None, ""):
            total_general += float(total)
        filas_tabla.append([
            Paragraph(str(fila.get("rg")) if fila.get("rg") not in (None, "") else "", estilo_celda),
            Paragraph(fila.get("codigo") or "", estilo_celda),
            Paragraph(fila.get("descripcion") or "", estilo_celda),
            Paragraph(str(fila.get("cantidad")) if fila.get("cantidad") not in (None, "") else "", estilo_celda),
            Paragraph(_formatear_moneda(fila.get("precio_unitario")), estilo_celda),
            Paragraph(_formatear_moneda(total), estilo_celda),
        ])

    tabla = Table(filas_tabla, colWidths=_anchos_columnas(borde_derecho), repeatRows=1)
    tabla.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("ALIGN", (0, 0), (1, -1), "CENTER"),
        ("ALIGN", (3, 0), (-1, -1), "CENTER"),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
    ]))

    estilo_total = ParagraphStyle("total", parent=estilos["Normal"], fontName="Helvetica-Bold", fontSize=9)
    estilo_normal_9 = ParagraphStyle("normal9", parent=estilos["Normal"], fontName="Helvetica", fontSize=9)
    estilo_firma = ParagraphStyle("firma", parent=estilos["Normal"], fontName="Helvetica-Bold", fontSize=9, alignment=1)

    story = [
        NextPageTemplate("continuacion"),
        tabla,
        Spacer(1, 10),
        Paragraph(f"Total: $ {_formatear_moneda(total_general)}", estilo_total),
        Paragraph(_monto_en_palabras(total_general), estilo_normal_9),
        Spacer(1, 24),
        Paragraph("FIRMA Y SELLO DEL PROPONENTE", estilo_firma),
    ]
    doc.build(story)
