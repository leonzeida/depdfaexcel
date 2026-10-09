#!/usr/bin/env python3
"""Genera una versión EDITADA del PDF F.41 original: mantiene el
membrete, el texto legal y la firma tal cual están en el PDF que se
subió, y completa la tabla de ítems con los precios que cargó el
usuario.

Dos caminos, elegidos automáticamente según qué tan distinta quedó la
lista de renglones de la original (ver _filas_coinciden_con_original):

1. SUPERPOSICIÓN (la mayoría de los casos reales: se cargan los mismos
   renglones del expediente y solo se completa el precio unitario). Se
   usan las páginas REALES del PDF original -- se les superpone el
   precio/total encima de los "$.........." en blanco que ya trae el
   formulario, sin rehacer la tabla. Esto garantiza que la cantidad de
   páginas y en qué página cae cada renglón sean EXACTAMENTE iguales al
   original, porque son, literalmente, las mismas páginas (a diferencia
   de reconstruir la tabla de cero, que nunca termina de replicar
   perfecto cómo pagina el sistema del gobierno -- confirmado comparando
   varios expedientes reales letra por letra).
2. RECONSTRUCCIÓN (si se agregó, sacó o reordenó algún renglón, o se
   editó el código/descripción/cantidad): ahí no alcanza con superponer,
   hay que rearmar la tabla entera. El membrete y el texto legal de la
   página 1 (y el bloque corto que se repite en cada página) se recortan
   del PDF original como IMÁGENES y se reusan tal cual -- no se
   reconstruyen a mano -- para garantizar que se vean exactamente
   iguales. Solo la tabla de ítems y el total final se dibujan de nuevo.

Se mantiene separado de f41_a_excel.py (que es de extracción/Excel).
"""

import io
import re
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import pdfplumber
from pypdf import PdfReader, PdfWriter
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas as canvas_module
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

from f41_a_excel import (
    COL_CANT_MAX,
    COL_CODIGO_MAX,
    COL_DESC_MAX,
    COL_RG_MAX,
    LEFT_BORDER_X,
    bandas_de_filas,
    extraer_pdf_completo,
)

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


FUENTE_TABLA = "Helvetica"
TAMANO_TABLA = 8
PADDING_CELDA = 3  # mismo valor que LEFTPADDING/RIGHTPADDING del TableStyle
# Interlineado real de "Descripción" en el F.41 original, medido letra por
# letra con pdfplumber: la altura de cada renglón de la tabla es, pixel a
# pixel, (cantidad de líneas que ocupa la descripción) x 11.25 -- sin
# ningún padding extra arriba/abajo (ver PADDING_VERTICAL_FILA). Con el
# valor viejo (leading 9.5 + padding 7+7) una descripción larga ocupaba
# notoriamente más alto que en el original, así que entraban menos
# renglones por página y sobraba una página entera (confirmado: mismo PDF,
# 4 páginas en el original contra 5 en el editado).
LEADING_TABLA = 11.25


def _texto_con_puntos(prefijo: str, ancho_objetivo: float, fuente: str, tamano: float) -> str:
    """`prefijo` + puntos de relleno hasta completar `ancho_objetivo`
    (medido con la fuente real, no un conteo fijo de caracteres). Si
    `prefijo` ya ocupa más que el ancho objetivo, se devuelve tal cual,
    sin puntos de más."""
    ancho_prefijo = stringWidth(prefijo, fuente, tamano)
    ancho_punto = stringWidth(".", fuente, tamano)
    cantidad_puntos = max(0, int((ancho_objetivo - ancho_prefijo) / ancho_punto))
    return prefijo + "." * cantidad_puntos


def _celda_precio(valor, ancho_columna: float) -> str:
    """Arma el texto de una celda de precio tal como lo hace el F.41
    original: "$" seguido de puntos de relleno hasta el ancho real de la
    columna (medido con la fuente real, no un conteo fijo de caracteres).
    Si hay un valor cargado, se escribe arriba de esa misma línea de
    puntos ("$ 1.234,56 ..........") en vez de en una celda aparte, para
    que se vea como un formulario completado a mano, no un documento
    distinto."""
    prefijo = "$" if valor in (None, "") else f"$ {_formatear_moneda(valor)} "
    return _texto_con_puntos(prefijo, ancho_columna - 2 * PADDING_CELDA, FUENTE_TABLA, TAMANO_TABLA)


def _medir_linea_visible(page, primera_palabra: str):
    """Busca la primera palabra == `primera_palabra` y mide el ancho
    visible de toda esa línea de texto (de la primera letra a la
    última), SIN contar los espacios en blanco de relleno que usa el
    PDF original para centrar el texto a mano (si se incluyeran, parecería
    que el texto arranca pegado al margen izquierdo cuando en realidad
    está centrado más a la derecha). Devuelve None si no la encuentra."""
    objetivo = next((w for w in page.extract_words() if w["text"] == primera_palabra), None)
    if not objetivo:
        return None
    chars_linea = [c for c in page.chars if abs(c["top"] - objetivo["top"]) < 1 and c["text"] != " "]
    if not chars_linea:
        return None
    return max(c["x1"] for c in chars_linea) - min(c["x0"] for c in chars_linea)


def _medir_linea_arriba_de(page, palabra_debajo: str, distancia_min=5, distancia_max=20):
    """Mide el ancho visible de la línea de puntos que está arriba de la
    palabra `palabra_debajo` (caso "FIRMA", con la línea para la firma
    arriba del texto "FIRMA Y SELLO DEL PROPONENTE"). Devuelve None si no
    la encuentra."""
    objetivo = next((w for w in page.extract_words() if w["text"] == palabra_debajo), None)
    if not objetivo:
        return None
    chars_arriba = [
        c for c in page.chars
        if distancia_min <= objetivo["top"] - c["top"] <= distancia_max and c["text"] not in (" ", "")
    ]
    if not chars_arriba:
        return None
    return max(c["x1"] for c in chars_arriba) - min(c["x0"] for c in chars_arriba)


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

        # El original no tiene ninguna línea vertical interna (confirmado
        # celda por celda): "Precio unitario"/"Total" son literalmente
        # "$" + puntos de relleno como texto, no columnas con borde. Se
        # detecta dónde arranca cada "$" en una fila de ítems real (no en
        # el encabezado) para poder replicar el mismo ancho exacto.
        x_precio_unitario = None
        x_total = None
        if len(bandas_p1) > 1:
            top_fila, bottom_fila = bandas_p1[1]
            dolares = sorted(
                c["x0"] for c in p1.chars
                if top_fila - 1 <= c["top"] <= bottom_fila + 1 and c["text"] == "$"
            )
            if len(dolares) >= 2:
                x_precio_unitario, x_total = dolares[0], dolares[1]
        if x_precio_unitario is None or x_total is None:
            x_precio_unitario = COL_CANT_MAX + 2
            x_total = COL_CANT_MAX + (borde_derecho - COL_CANT_MAX) / 2

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

        # "SON PESOS"/"FIRMA Y SELLO" solo aparecen en la ÚLTIMA página del
        # original (después de agotar los ítems), no en la página 1 -- en
        # un PDF de una sola página son la misma página. Se mide el ancho
        # visible real de esas líneas (sin los espacios de relleno que usa
        # el original para centrarlas a mano) para poder centrar lo mismo.
        p_ultima = pdf.pages[-1]
        ancho_son_pesos = _medir_linea_visible(p_ultima, "SON")
        ancho_firma = _medir_linea_arriba_de(p_ultima, "FIRMA")
        # Ancho visible de "Total: $..............": en el original NO
        # ocupa toda la fila, arranca bastante más a la derecha (recién
        # después de la mitad) y el relleno de puntos llega hasta el
        # borde -- medir solo este tramo (no todo el ancho de la tabla)
        # es lo que hace que, alineado a la derecha, el texto quede
        # corrido hacia la derecha en vez de pegado al margen izquierdo.
        ancho_total_linea = _medir_linea_visible(p_ultima, "Total:")

    return {
        "ancho": ancho,
        "alto": alto,
        "fin_encabezado_pagina1": fin_encabezado_pagina1,
        "fin_mini_encabezado": fin_mini_encabezado,
        "imagen_pagina1": imagen_pagina1,
        "imagen_mini": imagen_mini,
        "borde_derecho": borde_derecho,
        "x_precio_unitario": x_precio_unitario,
        "x_total": x_total,
        "ancho_son_pesos": ancho_son_pesos or 260,
        "ancho_firma": ancho_firma or 200,
        "ancho_total_linea": ancho_total_linea or 265,
        "pie_pagina": pie_pagina,
    }


def _anchos_columnas(info: dict) -> list:
    """Reusa los mismos límites de columna que ya usa la extracción
    (f41_a_excel.py) para Renglón/Código/Descripción, y las posiciones
    reales de "$" detectadas en _analizar_pdf_original para Cantidad/
    Precio unitario/Total, para que la tabla nueva quede alineada en la
    misma posición horizontal que la original."""
    return [
        COL_RG_MAX - LEFT_BORDER_X,
        COL_CODIGO_MAX - COL_RG_MAX,
        COL_DESC_MAX - COL_CODIGO_MAX,
        info["x_precio_unitario"] - COL_DESC_MAX,
        info["x_total"] - info["x_precio_unitario"],
        info["borde_derecho"] - info["x_total"],
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


def _texto(valor) -> str:
    return "" if valor is None else str(valor).strip()


def _numeros_iguales(a, b) -> bool:
    try:
        return float(a) == float(b)
    except (TypeError, ValueError):
        return _texto(a) == _texto(b)


def _filas_coinciden_con_original(filas: list, items_originales: list) -> bool:
    """True si `filas` (lo que arma la grilla, ya editado por el usuario)
    tiene exactamente los mismos renglones que `items_originales` (lo que
    se extrajo del PDF subido), en el mismo orden, con el mismo código/
    descripción/cantidad -- es decir, lo único que cambió fue completar
    precio_unitario/total. Es la condición para poder usar el método de
    superposición (ver módulo): si se agregó, sacó o reordenó algún
    renglón, o se tocó el código/descripción/cantidad, ya no es seguro
    superponer sobre las páginas originales (esos datos quedarían
    desactualizados en el PDF, que no los vuelve a dibujar)."""
    if len(filas) != len(items_originales):
        return False
    for fila, original in zip(filas, items_originales):
        if not isinstance(fila, dict):
            return False
        if _texto(fila.get("rg")) != _texto(original["rg"]):
            return False
        if _texto(fila.get("codigo")) != _texto(original["codigo"]):
            return False
        if _texto(fila.get("descripcion")) != _texto(original["descripcion"]):
            return False
        if not _numeros_iguales(fila.get("cantidad"), original["cantidad"]):
            return False
    return True


def _mapa_precios_original(pdf_path: Path) -> list:
    """Recorre el PDF original página por página y arma, para cada una,
    dónde hay que superponer qué: la posición real (x, top, bottom -- en
    el sistema de coordenadas de pdfplumber, origen arriba-izquierda) del
    "$" de Precio unitario y de Total de cada renglón, más la del "$" de
    la línea "Total:" final si esa página la tiene. Se basa en ubicar los
    caracteres "$" reales dentro de cada banda de fila (mismo criterio
    que _analizar_pdf_original usa para una sola fila de muestra, acá
    generalizado a todas), así el overlay queda pegado exactamente donde
    el formulario original tiene sus "$.........." en blanco."""
    paginas = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            palabras = page.extract_words()
            celdas = []
            total = None
            for top, bottom in bandas_de_filas(page):
                en_banda = [w for w in palabras if top - 0.5 <= w["top"] < bottom - 0.5]
                rg_texto = "".join(w["text"] for w in en_banda if w["x0"] < COL_RG_MAX)
                dolares = sorted(
                    (c for c in page.chars if top - 1 <= c["top"] <= bottom + 1 and c["text"] == "$"),
                    key=lambda c: c["x0"],
                )
                if rg_texto.isdigit() and len(dolares) >= 2:
                    celdas.append(
                        {
                            "rg": rg_texto,
                            "precio": dolares[0],
                            "total": dolares[1],
                        }
                    )
                elif not rg_texto.isdigit() and dolares and any(
                    w["text"] == "Total:" for w in en_banda
                ):
                    total = dolares[0]
            paginas.append(
                {
                    "ancho": page.width,
                    "alto": page.height,
                    "celdas": celdas,
                    "total": total,
                }
            )
    return paginas


def _generar_overlay_precios(pdf_original_path: Path, filas: list, info: dict, salida: Path):
    """Genera el F.41 editado superponiendo los precios sobre las páginas
    REALES del original (ver criterio en el docstring del módulo). Por
    cada "$.........." en blanco del original: tapa esa franja con un
    rectángulo blanco (mismo alto que la fila, para cubrir también los
    puntos de relleno existentes) y escribe encima el precio editado con
    el mismo criterio visual que usa la reconstrucción (_celda_precio)."""
    anchos_columnas = _anchos_columnas(info)
    ancho_precio_unitario, ancho_total = anchos_columnas[4], anchos_columnas[5]
    borde_derecho = info["borde_derecho"]

    items_por_rg = {_texto(f.get("rg")): f for f in filas}

    total_general = 0.0
    for fila in filas:
        total = fila.get("total")
        if total not in (None, ""):
            total_general += float(total)

    mapa = _mapa_precios_original(pdf_original_path)
    buffer_overlay = io.BytesIO()
    cv = canvas_module.Canvas(buffer_overlay)
    for pagina in mapa:
        cv.setPageSize((pagina["ancho"], pagina["alto"]))
        alto = pagina["alto"]

        def _tapar_y_escribir(x0, top, bottom, ancho_columna, texto, negrita=False):
            # +0.3/-0.3: un pelo de margen para cubrir también el borde
            # superior/inferior de los puntos de relleno sin invadir la
            # fila de al lado.
            cv.setFillColorRGB(1, 1, 1)
            cv.rect(x0 - 1, alto - bottom - 1, ancho_columna + 1, (bottom - top) + 2, fill=1, stroke=0)
            cv.setFillColorRGB(0, 0, 0)
            cv.setFont("Helvetica-Bold" if negrita else FUENTE_TABLA, TAMANO_TABLA)
            cv.drawString(x0, alto - bottom, texto)

        for celda in pagina["celdas"]:
            fila = items_por_rg.get(celda["rg"])
            if fila is None:
                continue
            d_precio, d_total = celda["precio"], celda["total"]
            _tapar_y_escribir(
                d_precio["x0"], d_precio["top"], d_precio["bottom"], ancho_precio_unitario,
                _celda_precio(fila.get("precio_unitario"), ancho_precio_unitario),
            )
            _tapar_y_escribir(
                d_total["x0"], d_total["top"], d_total["bottom"], ancho_total,
                _celda_precio(fila.get("total"), ancho_total),
            )

        if pagina["total"] is not None:
            d = pagina["total"]
            # El "Total:" en sí ya está impreso en el original (parte del
            # membrete/texto fijo que no se toca) -- acá solo se
            # superpone el "$.........." que sigue, igual que en cada
            # renglón de ítem. Sin esto quedaba "Total: Total: $..." por
            # duplicado.
            ancho_total_general = borde_derecho - d["x0"]
            _tapar_y_escribir(
                d["x0"], d["top"], d["bottom"], ancho_total_general,
                _celda_precio(total_general, ancho_total_general), negrita=True,
            )

        cv.showPage()
    cv.save()
    buffer_overlay.seek(0)

    lector_original = PdfReader(str(pdf_original_path))
    lector_overlay = PdfReader(buffer_overlay)
    writer = PdfWriter()
    for pagina_original, pagina_overlay in zip(lector_original.pages, lector_overlay.pages):
        pagina_original.merge_page(pagina_overlay)
        writer.add_page(pagina_original)

    with open(salida, "wb") as f:
        writer.write(f)


def _generar_reconstruccion(pdf_original_path: Path, filas: list, salida: Path):
    """Reconstruye la tabla de ítems de cero (ver criterio en el
    docstring del módulo: se usa cuando los renglones ya no coinciden con
    el original). `filas`: lista de dicts con "rg", "codigo",
    "descripcion", "cantidad", "precio_unitario", "total" (los valores de
    precio/total ya vienen resueltos del frontend, no se recalculan acá).
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
    estilo_celda = ParagraphStyle("celda", parent=estilos["Normal"], fontName=FUENTE_TABLA, fontSize=TAMANO_TABLA, leading=LEADING_TABLA)

    anchos_columnas = _anchos_columnas(info)
    ancho_precio_unitario, ancho_total = anchos_columnas[4], anchos_columnas[5]

    # El encabezado ("Rg Código Descripción Cant. P.Unit. Total") y las
    # filas de datos van como texto plano, no Paragraph: en el original
    # no hay wrap ni bordes internos, así que no hace falta ese motor de
    # texto más que para Descripción (la única columna que sí necesita
    # ajustar el texto a su ancho).
    encabezados = ["Rg", "Código", "Descripción", "Cant.", "P.Unit.", "Total"]
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
            _celda_precio(fila.get("precio_unitario"), ancho_precio_unitario),
            _celda_precio(total, ancho_total),
        ])

    # El "Total" general es, en el original, el último renglón de la
    # propia tabla (confirmado: bandas_de_filas() lo detecta como una
    # banda más, con el mismo borde que encierra toda la tabla), no un
    # párrafo aparte debajo -- con una línea arriba separándolo del
    # último ítem, el texto pegado al borde derecho (con el mismo
    # relleno de puntos que las celdas de precio) y el resto del
    # renglón en blanco. El relleno se calcula con el ancho MEDIDO de la
    # línea real (no con el ancho completo de la tabla): el original no
    # arranca "Total:" pegado al margen izquierdo, arranca bastante más a
    # la derecha -- como el párrafo queda alineado a la derecha, usar un
    # ancho más angosto es lo que corre todo el bloque hacia la derecha
    # en vez de ocupar toda la fila.
    fila_total_indice = len(filas_tabla)
    texto_total = _texto_con_puntos(
        f"Total: $ {_formatear_moneda(total_general)} ",
        info["ancho_total_linea"], "Helvetica-Bold", 8,
    )
    # El contenido va en la primera columna: con SPAN, reportlab arma la
    # celda fusionada a partir del contenido de la celda de más arriba a
    # la izquierda del rango (acá, columna 0) e ignora lo que haya en las
    # demás columnas fusionadas -- si el texto se pone en la última
    # columna, como en cualquier otra fila, no se ve.
    filas_tabla.append([texto_total, "", "", "", "", ""])

    # En el original los renglones de datos no tienen padding vertical
    # propio -- se separan solo por el interlineado de LEADING_TABLA (de
    # ahí que la altura real de cada fila sea exactamente líneas x 11.25,
    # confirmado letra por letra). Meterle padding acá además del
    # interlineado es lo que inflaba la altura de cada fila por encima del
    # original.
    PADDING_VERTICAL_FILA = 0

    tabla = Table(filas_tabla, colWidths=anchos_columnas, repeatRows=1)
    tabla.setStyle(TableStyle([
        # Ni líneas verticales internas, ni línea debajo del encabezado de
        # columnas, ni negrita ahí -- confirmado letra por letra en el PDF
        # original (no solo mirando las coordenadas): "Rg Código
        # Descripción..." usa la misma fuente Helvetica común que las
        # filas de datos, y en toda la zona de la tabla las únicas líneas
        # horizontales son la de arriba de "DETALLE DE ITEMS" y la del
        # final de la tabla -- ninguna entre el encabezado y el primer
        # renglón, ni entre renglones. O sea: el encabezado de columnas es
        # una fila más, sin ningún estilo especial salvo la alineación.
        ("BOX", (0, 0), (-1, -1), 0.5, colors.black),
        ("FONTNAME", (0, 0), (-1, -1), FUENTE_TABLA),
        ("FONTSIZE", (0, 0), (-1, -1), TAMANO_TABLA),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("ALIGN", (0, 0), (-1, -1), "LEFT"),
        # "P.Unit."/"Total" del encabezado van a la derecha (igual que el
        # original); el resto -- incluidas esas mismas columnas en las
        # filas de datos, donde el "$...puntos" ya arranca siempre a la
        # izquierda -- se queda alineado a la izquierda.
        ("ALIGN", (4, 0), (5, 0), "RIGHT"),
        ("TOPPADDING", (0, 0), (-1, -1), PADDING_VERTICAL_FILA),
        ("BOTTOMPADDING", (0, 0), (-1, -1), PADDING_VERTICAL_FILA),
        ("LEFTPADDING", (0, 0), (-1, -1), PADDING_CELDA),
        ("RIGHTPADDING", (0, 0), (-1, -1), PADDING_CELDA),
        # El renglón de "Total" es el último de la tabla: todas las
        # columnas se fusionan en una sola celda, con el texto pegado al
        # borde derecho (relleno de puntos ya incluido en texto_total) y
        # una línea arriba separándolo del último ítem -- igual que el
        # original.
        ("SPAN", (0, fila_total_indice), (-1, fila_total_indice)),
        ("FONTNAME", (0, fila_total_indice), (-1, fila_total_indice), "Helvetica-Bold"),
        ("FONTSIZE", (0, fila_total_indice), (-1, fila_total_indice), 8),
        ("ALIGN", (0, fila_total_indice), (-1, fila_total_indice), "RIGHT"),
        ("LINEABOVE", (0, fila_total_indice), (-1, fila_total_indice), 0.5, colors.black),
    ]))

    # "SON PESOS..." y "FIRMA Y SELLO DEL PROPONENTE" van centradas (en
    # el medio de la tabla), confirmado midiendo solo los caracteres
    # VISIBLES de cada línea en el original -- a primera vista parecen
    # arrancar pegadas al margen izquierdo, pero eso es porque el PDF
    # original las centra "a mano" con un montón de espacios en blanco
    # antes del texto, que si se cuentan como parte de la línea hacen
    # parecer que arranca en el margen. A pedido del usuario, "SON
    # PESOS" queda en blanco (con su línea de puntos, sin completar el
    # monto en palabras) -- igual que "FIRMA Y SELLO", es una línea para
    # completar a mano, no un dato que haya que resolver acá.
    estilo_son_pesos = ParagraphStyle("sonpesos", parent=estilos["Normal"], fontName="Helvetica-Bold", fontSize=9, alignment=1)
    estilo_firma_puntos = ParagraphStyle("firmapuntos", parent=estilos["Normal"], fontName="Helvetica-Bold", fontSize=9, alignment=1)
    estilo_firma = ParagraphStyle("firma", parent=estilos["Normal"], fontName="Helvetica-Bold", fontSize=9, alignment=1)

    texto_son_pesos = _texto_con_puntos("SON PESOS: ", info["ancho_son_pesos"], "Helvetica-Bold", 9)
    texto_firma_puntos = "." * max(1, int(info["ancho_firma"] / stringWidth(".", "Helvetica-Bold", 9)))

    story = [
        NextPageTemplate("continuacion"),
        tabla,
        Spacer(1, 24),
        Paragraph(texto_son_pesos, estilo_son_pesos),
        Spacer(1, 30),
        Paragraph(texto_firma_puntos, estilo_firma_puntos),
        Paragraph("FIRMA Y SELLO DEL PROPONENTE", estilo_firma),
    ]
    doc.build(story)


def generar_f41_editado(pdf_original_path: Path, filas: list, salida: Path):
    """`filas`: lista de dicts con "rg", "codigo", "descripcion",
    "cantidad", "precio_unitario", "total" (los valores de precio/total
    ya vienen resueltos del frontend, no se recalculan acá). Elige entre
    superposición (preferido: paginación idéntica al original) y
    reconstrucción (respaldo, si la lista de renglones ya no coincide con
    la original) -- ver criterio completo en el docstring del módulo."""
    try:
        items_originales = extraer_pdf_completo(pdf_original_path)
    except Exception:
        items_originales = None

    if items_originales and _filas_coinciden_con_original(filas, items_originales):
        info = _analizar_pdf_original(pdf_original_path)
        try:
            _generar_overlay_precios(pdf_original_path, filas, info, salida)
            return
        except Exception:
            # Si algo falla superponiendo (PDF original con una
            # estructura rara que no se pudo prever), no se pierde la
            # generación entera: se cae al método de reconstrucción, que
            # es más tolerante porque no depende de encontrar posiciones
            # exactas en el original.
            pass

    _generar_reconstruccion(pdf_original_path, filas, salida)
