#!/usr/bin/env python3
"""Genera el PDF del "Pedido de Cotización" que Zeid Medical les manda a
SUS PROPIOS proveedores (no es el F.41 del Gobierno de Formosa: es un
documento propio de la empresa, armado a partir de una lista de ítems que
el usuario edita libremente en la web — agrega, borra y reordena
renglones, y carga el precio unitario a su gusto).

Se mantiene separado de f41_a_excel.py a propósito: ese módulo genera
Excel, este genera PDF, son formatos y librerías distintas (mismo
criterio que ya separa db_referencia.py).
"""

from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

from num2words import num2words
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

ZONA_HORARIA_ARGENTINA = ZoneInfo("America/Argentina/Buenos_Aires")

# Anchos de columna (en mm), sumando el ancho útil de una A4 con 15mm de
# margen a cada lado (210 - 15 - 15 = 180mm).
ANCHOS_COLUMNAS_MM = (12, 25, 68, 15, 30, 30)


def _formatear_moneda(valor) -> str:
    """1234567.89 -> "1.234.567,89" (separador de miles con punto, coma
    decimal, como en toda la planilla de Excel del comparador)."""
    if valor is None or valor == "":
        return ""
    texto = f"{float(valor):,.2f}"
    # f-string da "1,234,567.89" (estilo US); se invierten los separadores.
    return texto.replace(",", "_").replace(".", ",").replace("_", ".")


def _monto_en_palabras(valor: float) -> str:
    centavos_totales = round(float(valor) * 100)
    entero, centavos = divmod(centavos_totales, 100)
    return f"SON PESOS: {num2words(entero, lang='es').upper()} PESOS CON {centavos:02d}/100.-"


def escribir_pedido_cotizacion_pdf(filas: list, salida: Path):
    """`filas`: lista de dicts con "rg", "codigo", "descripcion",
    "cantidad", "precio_unitario", "total". "precio_unitario"/"total" son
    los valores ya resueltos que manda el frontend (el cálculo automático
    de Cantidad×Precio unitario, o lo que el usuario haya pisado a mano
    encima) — acá no se recalculan, se confía en lo que llega.

    Genera un PDF propio de Zeid Medical (no el F.41 del gobierno) con
    esa tabla, el total general (suma de todos los "total") y el monto
    final escrito en palabras, estilo "SON PESOS: ... .-".
    """
    estilos = getSampleStyleSheet()
    estilo_celda = ParagraphStyle("celda", parent=estilos["Normal"], fontSize=9, leading=11)

    doc = SimpleDocTemplate(
        str(salida),
        pagesize=A4,
        topMargin=18 * mm,
        bottomMargin=18 * mm,
        leftMargin=15 * mm,
        rightMargin=15 * mm,
        title="Pedido de Cotización - Zeid Medical",
    )

    elementos = [
        Paragraph("ZEID MEDICAL S.R.L.", ParagraphStyle(
            "titulo", parent=estilos["Title"], fontSize=16, spaceAfter=2,
        )),
        Paragraph("Pedido de Cotización", estilos["Heading2"]),
        Paragraph(
            "Fecha: " + datetime.now(ZONA_HORARIA_ARGENTINA).strftime("%d/%m/%Y"),
            estilos["Normal"],
        ),
        Spacer(1, 8 * mm),
    ]

    encabezados = ["Rg.", "Código", "Descripción", "Cant.", "Precio unitario", "Total"]
    filas_tabla = [encabezados]
    total_general = 0.0
    for fila in filas:
        total = fila.get("total")
        if total not in (None, ""):
            total_general += float(total)
        filas_tabla.append([
            str(fila.get("rg")) if fila.get("rg") not in (None, "") else "",
            fila.get("codigo") or "",
            Paragraph(fila.get("descripcion") or "", estilo_celda),
            str(fila.get("cantidad")) if fila.get("cantidad") not in (None, "") else "",
            _formatear_moneda(fila.get("precio_unitario")),
            _formatear_moneda(total),
        ])

    tabla = Table(
        filas_tabla,
        colWidths=[a * mm for a in ANCHOS_COLUMNAS_MM],
        repeatRows=1,
    )
    tabla.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#d4ea6b")),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#999999")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("ALIGN", (0, 0), (1, -1), "CENTER"),
        ("ALIGN", (3, 0), (-1, -1), "CENTER"),
        ("ALIGN", (4, 0), (-1, -1), "RIGHT"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
    ]))
    elementos.append(tabla)

    elementos.append(Spacer(1, 8 * mm))
    elementos.append(Paragraph(
        f"Total: $ {_formatear_moneda(total_general)}",
        ParagraphStyle("total", parent=estilos["Normal"], fontName="Helvetica-Bold", fontSize=11),
    ))
    elementos.append(Paragraph(_monto_en_palabras(total_general), estilos["Normal"]))

    doc.build(elementos)
