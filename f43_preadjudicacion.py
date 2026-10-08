#!/usr/bin/env python3
"""Extrae los items adjudicados de un Acta de Preadjudicacion F.43
(Gobierno de la Provincia de Formosa) - el documento que publica, por
cada renglon de una licitacion, que proveedor gano y a que precio.

A diferencia del F.41 (f41_a_excel.py), este PDF no tiene lineas
verticales ni encabezados de columna impresos: las columnas son
puramente posicionales, y el proveedor no es una columna de la tabla
sino un bloque de texto ("Firma Adjudicada: ...") que agrupa 1 o mas
renglones. Ver el modulo f41_a_excel.py para la logica del otro tipo de
PDF (no se puede reusar aca: ese parser depende de bordes verticales que
este documento no tiene).
"""

import re
from pathlib import Path

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

from f41_a_excel import FMT_MONEDA_ARS

# Coordenadas (en puntos) de las columnas de la tabla del F.43. A
# diferencia del F.41 son puramente posicionales (no hay lineas
# verticales en el PDF que las delimiten) pero son fijas entre
# expedientes porque el formulario es una plantilla oficial.
COL_RG_MAX = 96.4
COL_CODIGO_MAX = 155.9
COL_DESC_MAX = 311.8
COL_MODALIDAD_MAX = 368.5
COL_CANT_MAX = 411.0
COL_PUNIT_MAX = 496.1

# Texto repetido en el encabezado/pie de cada pagina (no es parte de la
# tabla ni del parrafo inicial) - se descarta para que no contamine la
# extraccion de renglones cuando una fila queda partida entre paginas.
RE_BOILERPLATE = re.compile(
    r"^(F\.43|Provincia de Formosa|Unidad de Compras de Productos e Insumos Medicinales|"
    r"ACTA DE PREADJUDICACION|Contrataci[oó]n o Compra Directa|"
    r"ADQUISICI[OÓ]N DE PRODUCTOS E INSUMOS MEDICINALES EXPTE N[º°]|"
    r"Comprobante N[º°]:|Usuario .+ P[áa]gina \d+)",
)

RE_FIRMA_ADJUDICADA = re.compile(r"^Firma Adjudicada:\s*(.+?)\s*-\s*C\.U\.I\.T:", re.IGNORECASE)
RE_TOTAL_ADJUDICADO = re.compile(r"^Total Adjudicado a\b", re.IGNORECASE)
# Codigo de RUBRO (categoria, no de producto): ej. "20.00.00.A15". Se usa
# para detectar donde empieza el "relleno" del PROXIMO item (nombre de
# rubro que aparece antes de su numero de renglon) y no confundirlo con
# el codigo del item actual.
RE_CODIGO_RUBRO = re.compile(r"^\d{2}\.00\.00\.[A-Z]\d{2}$")
# Codigo de ITEM (producto): mismo esquema que F.41 (PREFIJOS_CODIGO,
# ej. "4.01.008.043").
RE_CODIGO_ITEM = re.compile(r"^\d\.\d{2}\.\d{3}\.\d{3}$")

# El prefijo de letra ("U-") queda fuera del grupo capturado a propósito,
# igual criterio que RE_EXPEDIENTE en f41_a_excel.py (que tampoco lo
# incluye) - así `encabezado["expediente"]` queda con el mismo formato en
# los dos modulos y se puede reusar `nombre_archivo_notas_pedido` del
# F.41 sin que duplique el prefijo.
RE_EXPEDIENTE_F43 = re.compile(r"EXPTE N[º°]\s*[A-Z]?-?(\d+/\d+)")
# A diferencia del F.41 (cuyo regex exige el formato "N°:") el F.43 no
# lleva los dos puntos despues de "N°" ("CONTRATACIÓN DIRECTA N° 0588/26").
RE_CONTRATACION_F43 = re.compile(r"CONTRATACI[ÓO]N DIRECTA N[º°]\s*:?\s*(\d+/\d+)")
# La fecha/hora de apertura viene en prosa, no como "APERTURA: dd/mm/aaaa"
# igual que en el F.41 (ej. "a los TREINTA (30) días del mes de
# SEPTIEMBRE de 2026, siendo las 09:55 horas").
RE_FECHA_F43 = re.compile(r"a los .+?\((\d{1,2})\)\s*d[ií]as del mes de (\w+) de (\d{4})", re.IGNORECASE)
RE_HORA_F43 = re.compile(r"siendo las (\d{1,2}:\d{2}) horas", re.IGNORECASE)
MESES_ES = {
    "ENERO": 1, "FEBRERO": 2, "MARZO": 3, "ABRIL": 4, "MAYO": 5, "JUNIO": 6,
    "JULIO": 7, "AGOSTO": 8, "SEPTIEMBRE": 9, "OCTUBRE": 10, "NOVIEMBRE": 11, "DICIEMBRE": 12,
}

# Palabras sueltas de la misma linea que su numero de renglon se agrupan
# por redondeo de la coordenada vertical ("top"); el margen de tolerancia
# evita que offsets de 0.1-0.2pt entre palabras de una misma fila visual
# (confirmado con datos reales: ocurre con las notas "Obs.:") las separe
# en dos lineas distintas.
TOL_LINEA = 2.0
# Salto vertical (en puntos) que separa un simple wrap de texto dentro de
# la misma celda de un nombre de rubro de mas de una linea empezando a
# aparecer (relleno del PROXIMO item, ver mas abajo). Se determino
# comparando la distribucion real de saltos entre lineas consecutivas en
# los PDF de ejemplo: los wraps dentro de una celda miden ~5-11pt, los
# saltos entre bloques (item a item, o item a marcador) miden ~17-24pt.
UMBRAL_SALTO_FILA = 14.0


def _agrupar_lineas(pdf_path: Path):
    """Todas las palabras del PDF, de todas las paginas, como un unico
    stream ordenado de arriba a abajo (con un offset vertical acumulado
    por pagina) agrupadas en lineas fisicas. Procesar como stream
    continuo - en vez de pagina por pagina como hace F.41 - es necesario
    porque un mismo renglon puede quedar partido entre el final de una
    pagina y el principio de la siguiente."""
    with pdfplumber.open(pdf_path) as pdf:
        palabras = []
        offset = 0.0
        for page in pdf.pages:
            for w in page.extract_words():
                w = dict(w)
                w["top"] += offset
                w["bottom"] += offset
                palabras.append(w)
            offset += page.height

    palabras.sort(key=lambda w: (w["top"], w["x0"]))
    lineas = []
    actual = []
    ref_top = None
    for w in palabras:
        if ref_top is None or w["top"] - ref_top <= TOL_LINEA:
            actual.append(w)
            if ref_top is None:
                ref_top = w["top"]
        else:
            lineas.append(actual)
            actual = [w]
            ref_top = w["top"]
    if actual:
        lineas.append(actual)

    resultado = []
    for linea in lineas:
        linea_ordenada = sorted(linea, key=lambda w: w["x0"])
        texto = " ".join(w["text"] for w in linea_ordenada)
        if RE_BOILERPLATE.match(texto):
            continue
        top_promedio = sum(w["top"] for w in linea_ordenada) / len(linea_ordenada)
        resultado.append((top_promedio, linea_ordenada))
    return resultado


def _parsear_moneda(texto: str):
    texto = texto.replace("$", "").strip()
    texto = texto.replace(".", "").replace(",", ".")
    try:
        return float(texto)
    except ValueError:
        return None


def _procesar_resto_de_linea(palabras, item_actual):
    """Busca el codigo de item (columna Codigo) en `palabras` si todavia
    no se encontro, y bucketiza el resto de las palabras en
    Descripcion/Cantidad/Precio unitario segun su posicion x. Sirve tanto
    para la linea ancla (la que trae el numero de renglon) como para las
    lineas de continuacion: en los PDF reales el codigo del item puede
    aparecer en cualquiera de las dos."""
    if not item_actual["codigo"]:
        codigo_candidato = "".join(
            w["text"] for w in palabras if COL_RG_MAX <= w["x0"] < COL_CODIGO_MAX
        )
        if RE_CODIGO_ITEM.match(codigo_candidato):
            item_actual["codigo"] = codigo_candidato

    for w in palabras:
        if COL_CODIGO_MAX <= w["x0"] < COL_DESC_MAX and w["x1"] <= COL_DESC_MAX:
            item_actual["_desc"].append(w)
        elif w["x0"] >= COL_CODIGO_MAX and COL_MODALIDAD_MAX < w["x1"] <= COL_CANT_MAX:
            item_actual["_cant"].append(w)
        elif COL_CANT_MAX < w["x1"] <= COL_PUNIT_MAX:
            item_actual["_punit"].append(w)


def _palabras_desc_de_linea(linea):
    return [w for w in linea if COL_CODIGO_MAX <= w["x0"] < COL_DESC_MAX and w["x1"] <= COL_DESC_MAX]


def extraer_adjudicaciones(pdf_path: Path) -> list:
    """Lista de items adjudicados: {renglon, codigo, descripcion,
    cantidad, precio_unitario, proveedor}, en el mismo orden en que
    aparecen en el PDF (agrupados por "Firma Adjudicada", no por numero
    de renglon).

    El documento no tiene una columna de proveedor: el nombre aparece
    como un bloque "Firma Adjudicada: <nombre> - C.U.I.T: <cuit>" que
    agrupa 1 o mas renglones siguientes, cerrado por "Total Adjudicado a
    ...". Tampoco hay lineas verticales que delimiten las filas (a
    diferencia del F.41), asi que cada renglon nuevo se ancla en la
    palabra de la columna Renglon que sea un digito puro.

    Antes de cada renglon aparece tambien el nombre de "rubro" (categoria
    del insumo, ej. "ANTIVIRAL.", "INSUMOS DE USO MEDICO Y DE
    LABORATORIO.") del PROXIMO item, que puede compartir linea con el
    comienzo real de su descripcion - por eso ese texto no se descarta:
    se junta en `buffer_pendiente` y se le agrega al item cuando se
    encuentra su ancla de renglon. El rubro queda como prefijo de la
    Descripcion final (mismo criterio que ya usa el F.41 para este tipo
    de items, que tampoco lo separa)."""
    lineas = _agrupar_lineas(pdf_path)

    items = []
    item_actual = None
    proveedor_actual = None
    # True hasta la primera "Firma Adjudicada:" - descarta el parrafo
    # narrativo de la pagina 0 (que puede traer numeros sueltos que de
    # otra forma podrian confundirse con un renglon).
    en_header_tabla = True
    # True mientras no se esta dentro del contenido ya confirmado de un
    # item (antes del primer renglon de una Firma Adjudicada, justo
    # despues de un Total Adjudicado, o mientras aparece el nombre de
    # rubro del PROXIMO item) - en ese estado, el texto de Descripcion
    # que se va viendo se junta en `buffer_pendiente` en vez de
    # agregarse al item actualmente abierto.
    modo_pendiente = True
    buffer_pendiente = []
    top_anterior = None

    def cerrar_item():
        nonlocal item_actual
        if item_actual is None:
            return
        desc_palabras = sorted(item_actual["_desc"], key=lambda w: (round(w["top"], 1), w["x0"]))
        descripcion = " ".join(w["text"] for w in desc_palabras).strip()
        if item_actual["_obs"]:
            descripcion = (descripcion + " (Obs.: " + " ".join(item_actual["_obs"]) + ")").strip()
        cant_texto = "".join(w["text"] for w in item_actual["_cant"])
        punit_texto = "".join(w["text"] for w in item_actual["_punit"]).replace("$", "")
        items.append(
            {
                "renglon": item_actual["renglon"],
                "codigo": item_actual["codigo"],
                "descripcion": descripcion,
                "cantidad": int(cant_texto) if cant_texto.isdigit() else cant_texto,
                "precio_unitario": _parsear_moneda(punit_texto) if punit_texto else None,
                "proveedor": item_actual["proveedor"],
            }
        )
        item_actual = None

    for top, linea in lineas:
        texto = " ".join(w["text"] for w in linea)
        gap = None if top_anterior is None else (top - top_anterior)
        top_anterior = top

        m_firma = RE_FIRMA_ADJUDICADA.match(texto)
        if m_firma:
            cerrar_item()
            proveedor_actual = m_firma.group(1).strip()
            en_header_tabla = False
            modo_pendiente = True
            buffer_pendiente = []
            continue

        if RE_TOTAL_ADJUDICADO.match(texto):
            cerrar_item()
            modo_pendiente = True
            buffer_pendiente = []
            continue

        rg_palabras = [w for w in linea if w["x0"] < COL_RG_MAX]
        rg_texto = "".join(w["text"] for w in rg_palabras)

        if rg_texto.isdigit() and not en_header_tabla:
            cerrar_item()
            resto = [w for w in linea if w["x0"] >= COL_RG_MAX]
            item_actual = {
                "renglon": int(rg_texto),
                "proveedor": proveedor_actual,
                "codigo": "",
                "_desc": list(buffer_pendiente),
                "_cant": [],
                "_punit": [],
                "_obs": [],
            }
            buffer_pendiente = []
            modo_pendiente = False
            _procesar_resto_de_linea(resto, item_actual)
            continue

        if en_header_tabla:
            continue

        if any(w["text"] == "Obs.:" for w in linea):
            if item_actual is not None and not modo_pendiente:
                nota = " ".join(w["text"] for w in linea if w["text"] != "Obs.:")
                item_actual["_obs"].append(nota)
            continue

        if modo_pendiente:
            buffer_pendiente.extend(_palabras_desc_de_linea(linea))
            continue

        if item_actual is None:
            continue

        codigo_candidato = "".join(
            w["text"] for w in linea if COL_RG_MAX <= w["x0"] < COL_CODIGO_MAX
        )
        if RE_CODIGO_RUBRO.match(codigo_candidato) and not item_actual["codigo"]:
            # Arranca el relleno (nombre de rubro) del PROXIMO item: lo
            # que haya de Descripcion en esta misma linea ya es parte de
            # ese proximo item, no del actual.
            modo_pendiente = True
            buffer_pendiente = _palabras_desc_de_linea(linea)
            continue

        # Linea ambigua (sin codigo de item ni de rubro reconocible):
        # puede ser continuacion real del item actual (wrap corto, salto
        # chico) o el nombre de rubro del PROXIMO item empezando a
        # aparecer en mas de una linea (salto grande).
        if not RE_CODIGO_ITEM.match(codigo_candidato) and gap is not None and gap > UMBRAL_SALTO_FILA:
            modo_pendiente = True
            buffer_pendiente = _palabras_desc_de_linea(linea)
            continue

        _procesar_resto_de_linea(linea, item_actual)

    cerrar_item()
    return items


def extraer_items_desiertos(pdf_path: Path) -> list:
    """Numeros de renglon que el acta declara "desiertos" (sin
    proveedor adjudicado). No son filas de la tabla: solo se mencionan
    en una frase narrativa de la pagina 0, con formato inconsistente
    entre expedientes (a veces un "ITEM" por vez con su propio motivo, a
    veces "ITEMS" seguido de una lista larga separada por comas) - el
    parser es best-effort y devuelve una lista vacia si no reconoce la
    frase, en vez de romper."""
    with pdfplumber.open(pdf_path) as pdf:
        texto0 = pdf.pages[0].extract_text() or ""
    texto0 = re.sub(r"\s+", " ", texto0)

    idx = texto0.lower().find("deja constancia que")
    if idx == -1:
        return []
    resto = texto0[idx + len("deja constancia que"):]
    fin = resto.find(".-")
    bloque = resto[:fin] if fin != -1 else resto[:500]

    numeros = set()
    recolectando = False
    for tok in bloque.split(" "):
        limpio = tok.strip(",.")
        if limpio.upper() in ("ITEM", "ITEMS"):
            recolectando = True
            continue
        if recolectando:
            if limpio.isdigit():
                numeros.add(int(limpio))
                continue
            if limpio.lower() == "y" or limpio == "":
                continue
            recolectando = False
    return sorted(numeros)


# Nombre tal cual aparece en "Firma Adjudicada: ZEID MEDICAL S.R.L. -
# C.U.I.T: ...". Se compara por substring en mayusculas (no exacto) para
# no depender de variantes menores de puntuacion/espacios.
NOMBRE_ZEID_MEDICAL = "ZEID MEDICAL"


def es_zeid_medical(proveedor: str) -> bool:
    return NOMBRE_ZEID_MEDICAL in (proveedor or "").upper()


def extraer_preadjudicacion_completa(pdf_path: Path) -> list:
    """Adjudicados + desiertos (estos ultimos con proveedor "DESIERTO" y
    el resto de los campos vacios, ya que el PDF no trae mas datos para
    ellos). Los adjudicados se devuelven en el mismo orden en que
    aparecen en el PDF (agrupados por proveedor/Firma Adjudicada, no por
    numero de renglon); los desiertos - que no tienen una posicion propia
    en la tabla, solo se mencionan en una frase aparte - se agregan al
    final, ordenados entre si por numero de renglon."""
    adjudicados = extraer_adjudicaciones(pdf_path)
    renglones_adjudicados = {it["renglon"] for it in adjudicados}

    desiertos = [
        {
            "renglon": renglon,
            "codigo": "",
            "descripcion": "",
            "cantidad": "",
            "precio_unitario": None,
            "proveedor": "DESIERTO",
        }
        for renglon in extraer_items_desiertos(pdf_path)
        if renglon not in renglones_adjudicados
    ]

    return adjudicados + desiertos


def extraer_encabezado_f43(pdf_path: Path) -> dict:
    """Ademas de expediente/contratacion, incluye apertura ("dd/mm/aaaa")
    y hora ("HH:MM") - con el mismo formato que usa `encabezado` en
    f41_a_excel.py - para poder reusar `escribir_notas_pedido` y
    `nombre_archivo_notas_pedido` de ese modulo al armar la Nota de
    Pedido de los renglones que ganó Zeid Medical."""
    with pdfplumber.open(pdf_path) as pdf:
        texto = pdf.pages[0].extract_text() or ""
    texto_una_linea = re.sub(r"\s+", " ", texto)

    m_exp = RE_EXPEDIENTE_F43.search(texto)
    m_con = RE_CONTRATACION_F43.search(texto)
    m_fecha = RE_FECHA_F43.search(texto_una_linea)
    m_hora = RE_HORA_F43.search(texto_una_linea)

    apertura = ""
    if m_fecha:
        dia, mes_nombre, anio = m_fecha.groups()
        mes_num = MESES_ES.get(mes_nombre.upper())
        if mes_num:
            apertura = f"{int(dia):02d}/{mes_num:02d}/{anio}"

    return {
        "expediente": m_exp.group(1) if m_exp else "",
        "contratacion": m_con.group(1) if m_con else "",
        "apertura": apertura,
        "hora": m_hora.group(1) if m_hora else "",
    }


def titulo_preadjudicacion(encabezado: dict) -> str:
    partes = ["Acta de Preadjudicación"]
    if encabezado.get("expediente"):
        partes.append(f"Expte. N° U-{encabezado['expediente']}")
    if encabezado.get("contratacion"):
        partes.append(f"Contratación Directa N° {encabezado['contratacion']}")
    return " — ".join(partes)


ENCABEZADOS_PREADJUDICACION = [
    "Renglon", "Codigo", "Descripcion", "Cantidad", "Precio", "Proveedor", "Costo", "%", "P. Venta",
]


def escribir_preadjudicacion_excel(titulo: str, filas: list, salida: Path):
    """Vuelca a un .xlsx la grilla de "Precios adjudicados" tal cual la
    ve el usuario (9 columnas). `filas` es una lista de listas de 9
    valores en el mismo orden que ENCABEZADOS_PREADJUDICACION."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Precios adjudicados"

    ws.merge_cells("A1:I1")
    ws["A1"] = titulo
    ws["A1"].font = Font(bold=True, size=13, color="000000")
    ws["A1"].alignment = Alignment(horizontal="center")
    ws["A1"].fill = PatternFill("solid", fgColor="D4EA6B")

    fuente_encabezado = Font(bold=True, color="000000")
    relleno_encabezado = PatternFill("solid", fgColor="D4EA6B")
    centrado = Alignment(horizontal="center")

    for col, titulo_columna in enumerate(ENCABEZADOS_PREADJUDICACION, start=1):
        c = ws.cell(row=2, column=col, value=titulo_columna)
        c.font = fuente_encabezado
        c.fill = relleno_encabezado
        c.alignment = centrado

    columnas_moneda = (5, 7, 9)  # Precio, Costo, P. Venta
    fila_excel = 3
    for fila in filas:
        fila = (list(fila) + [None] * 9)[:9]
        for col, valor in enumerate(fila, start=1):
            c = ws.cell(row=fila_excel, column=col, value=valor if valor != "" else None)
            c.alignment = Alignment(horizontal="left", wrap_text=True) if col == 3 else centrado
            if col in columnas_moneda:
                c.number_format = FMT_MONEDA_ARS
        fila_excel += 1

    anchos = {"A": 10, "B": 16, "C": 46, "D": 10, "E": 14, "F": 28, "G": 14, "H": 8, "I": 14}
    for letra, ancho in anchos.items():
        ws.column_dimensions[letra].width = ancho

    wb.save(salida)
