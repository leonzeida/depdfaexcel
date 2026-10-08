#!/usr/bin/env python3
"""Acceso a la base de datos de Zeid Medical: precios de referencia del
comparador de precios de proveedores, y el historial de licitaciones
(oferta -> resultado real -> costo real) que vincula el comparador (F.41)
con las Actas de Preadjudicación (F.43).

Se mantiene separado de f41_a_excel.py a propósito: ese módulo no debería
necesitar una base de datos para poder usarse por línea de comandos.

Requiere una variable de entorno:
- DATABASE_URL: connection string de Postgres (ej. la que da Neon), con
  el formato postgres://usuario:contraseña@host/db?sslmode=require.
"""

import os
import re
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import psycopg
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

from f41_a_excel import FMT_MONEDA_ARS, clave_item

ZONA_HORARIA_ARGENTINA = ZoneInfo("America/Argentina/Buenos_Aires")

_CREAR_TABLA = """
CREATE TABLE IF NOT EXISTS precios_referencia (
    codigo TEXT NOT NULL,
    descripcion TEXT NOT NULL,
    ultimo_precio NUMERIC(12,2),
    actualizado DATE,
    PRIMARY KEY (codigo, descripcion)
);
"""

# ADD COLUMN IF NOT EXISTS para no romper la tabla ya creada en producción
# (antes de agregar "porcentaje", el CREATE TABLE de arriba ya se había
# corrido ahí, así que un CREATE TABLE nuevo no le agrega la columna sola).
_AGREGAR_COLUMNA_PORCENTAJE = """
ALTER TABLE precios_referencia ADD COLUMN IF NOT EXISTS porcentaje NUMERIC(12,2);
"""

# Mismo criterio: se agrega después de que la tabla ya existía en
# producción sin esta columna.
_AGREGAR_COLUMNA_MEJOR_PROVEEDOR = """
ALTER TABLE precios_referencia ADD COLUMN IF NOT EXISTS mejor_proveedor TEXT;
"""

_UPSERT = """
INSERT INTO precios_referencia (codigo, descripcion, ultimo_precio, porcentaje, mejor_proveedor, actualizado)
VALUES (%s, %s, %s, %s, %s, %s)
ON CONFLICT (codigo, descripcion) DO UPDATE SET
    ultimo_precio = EXCLUDED.ultimo_precio,
    porcentaje = EXCLUDED.porcentaje,
    mejor_proveedor = EXCLUDED.mejor_proveedor,
    actualizado = EXCLUDED.actualizado;
"""

# Una fila por (item, licitacion), que se va completando en 3 etapas a
# medida que pasa el tiempo en vez de pisarse (a diferencia de
# precios_referencia, que es "la ultima foto" de cada item):
#   1. oferta (comparador F.41): que precio calculo Zeid, con que proveedor propio.
#   2. resultado (F.43): a quien se lo adjudicaron en la realidad y a que precio.
#   3. costo real (F.43, items ganados): cuanto salio comprarlo de verdad.
_CREAR_TABLA_HISTORIAL = """
CREATE TABLE IF NOT EXISTS historial_licitaciones (
    codigo TEXT NOT NULL,
    descripcion_normalizada TEXT NOT NULL,
    expediente TEXT NOT NULL,
    descripcion TEXT,
    contratacion TEXT,
    cantidad NUMERIC(12,2),
    precio_ofertado NUMERIC(12,2),
    porcentaje_ofertado NUMERIC(12,2),
    proveedor_elegido TEXT,
    fecha_oferta DATE,
    gano_zeid BOOLEAN,
    precio_adjudicado NUMERIC(12,2),
    proveedor_ganador TEXT,
    fecha_adjudicacion DATE,
    costo_real NUMERIC(12,2),
    precio_venta_real NUMERIC(12,2),
    fecha_costo DATE,
    PRIMARY KEY (codigo, descripcion_normalizada, expediente)
);
"""

# COALESCE(EXCLUDED.x, tabla.x) en vez de pisar directo: las etapas 1/2/3
# se guardan en momentos distintos (a veces días de diferencia) y cada
# upsert solo trae sus propias columnas - sin esto, guardar la etapa 2
# pisaría con NULL lo que ya se había guardado en la etapa 1, y viceversa.
_UPSERT_OFERTA = """
INSERT INTO historial_licitaciones (
    codigo, descripcion_normalizada, expediente, descripcion, contratacion, cantidad,
    precio_ofertado, porcentaje_ofertado, proveedor_elegido, fecha_oferta
) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
ON CONFLICT (codigo, descripcion_normalizada, expediente) DO UPDATE SET
    descripcion = COALESCE(EXCLUDED.descripcion, historial_licitaciones.descripcion),
    contratacion = COALESCE(EXCLUDED.contratacion, historial_licitaciones.contratacion),
    cantidad = COALESCE(EXCLUDED.cantidad, historial_licitaciones.cantidad),
    precio_ofertado = COALESCE(EXCLUDED.precio_ofertado, historial_licitaciones.precio_ofertado),
    porcentaje_ofertado = COALESCE(EXCLUDED.porcentaje_ofertado, historial_licitaciones.porcentaje_ofertado),
    proveedor_elegido = COALESCE(EXCLUDED.proveedor_elegido, historial_licitaciones.proveedor_elegido),
    fecha_oferta = COALESCE(EXCLUDED.fecha_oferta, historial_licitaciones.fecha_oferta);
"""

_UPSERT_ADJUDICACION = """
INSERT INTO historial_licitaciones (
    codigo, descripcion_normalizada, expediente, descripcion, contratacion, cantidad,
    gano_zeid, precio_adjudicado, proveedor_ganador, fecha_adjudicacion,
    costo_real, precio_venta_real, fecha_costo
) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
ON CONFLICT (codigo, descripcion_normalizada, expediente) DO UPDATE SET
    descripcion = COALESCE(EXCLUDED.descripcion, historial_licitaciones.descripcion),
    contratacion = COALESCE(EXCLUDED.contratacion, historial_licitaciones.contratacion),
    cantidad = COALESCE(EXCLUDED.cantidad, historial_licitaciones.cantidad),
    gano_zeid = COALESCE(EXCLUDED.gano_zeid, historial_licitaciones.gano_zeid),
    precio_adjudicado = COALESCE(EXCLUDED.precio_adjudicado, historial_licitaciones.precio_adjudicado),
    proveedor_ganador = COALESCE(EXCLUDED.proveedor_ganador, historial_licitaciones.proveedor_ganador),
    fecha_adjudicacion = COALESCE(EXCLUDED.fecha_adjudicacion, historial_licitaciones.fecha_adjudicacion),
    costo_real = COALESCE(EXCLUDED.costo_real, historial_licitaciones.costo_real),
    precio_venta_real = COALESCE(EXCLUDED.precio_venta_real, historial_licitaciones.precio_venta_real),
    fecha_costo = COALESCE(EXCLUDED.fecha_costo, historial_licitaciones.fecha_costo);
"""


class ErrorPreciosReferencia(Exception):
    """Cualquier problema al leer o guardar en la base de datos."""


def _conectar():
    url = os.environ.get("DATABASE_URL")
    if not url:
        raise ErrorPreciosReferencia("Falta configurar DATABASE_URL.")
    try:
        conn = psycopg.connect(url)
    except Exception as exc:
        raise ErrorPreciosReferencia(
            f"No se pudo conectar a la base de datos: {type(exc).__name__}: {exc!r}"
        ) from exc
    cur = conn.cursor()
    cur.execute(_CREAR_TABLA)
    cur.execute(_AGREGAR_COLUMNA_PORCENTAJE)
    cur.execute(_AGREGAR_COLUMNA_MEJOR_PROVEEDOR)
    cur.execute(_CREAR_TABLA_HISTORIAL)
    cur.close()
    conn.commit()
    return conn


def _normalizar_descripcion(descripcion: str) -> str:
    """Clave de match entre documentos: el mismo item puede venir con la
    descripcion formateada un poco distinto segun si se extrajo de un
    F.41 o de un F.43 (espacios, como se agrega la nota "Obs.:") -
    confirmado comparando el mismo expediente en los dos formatos. Sin
    normalizar, la etapa 1 (guardada desde el F.41) y la etapa 2
    (guardada desde el F.43) de un mismo item terminarian en filas
    distintas en vez de completar la misma."""
    texto = re.sub(r"\(Obs\.:.*?\)", "", descripcion or "", flags=re.IGNORECASE)
    texto = re.sub(r"\s+", " ", texto).strip().upper()
    return texto


def leer_precios_referencia() -> dict:
    """Devuelve {(codigo, descripcion): {"ultimo_precio": float|None,
    "porcentaje": float|None, "actualizado": str|None (fecha ISO),
    "mejor_proveedor": str|None}}."""
    conn = None
    try:
        conn = _conectar()
        cur = conn.cursor()
        cur.execute(
            "SELECT codigo, descripcion, ultimo_precio, porcentaje, actualizado, mejor_proveedor "
            "FROM precios_referencia;"
        )
        filas = cur.fetchall()
        cur.close()
    except ErrorPreciosReferencia:
        raise
    except Exception as exc:
        raise ErrorPreciosReferencia(
            f"No se pudo leer la tabla de precios de referencia: {type(exc).__name__}: {exc!r}"
        ) from exc
    finally:
        if conn is not None:
            conn.close()

    precios = {}
    for codigo, descripcion, ultimo_precio, porcentaje, actualizado, mejor_proveedor in filas:
        precios[clave_item(codigo, descripcion)] = {
            "ultimo_precio": float(ultimo_precio) if ultimo_precio is not None else None,
            "porcentaje": float(porcentaje) if porcentaje is not None else None,
            "actualizado": actualizado.isoformat() if actualizado is not None else None,
            "mejor_proveedor": mejor_proveedor,
        }
    return precios


def guardar_precios_referencia(items: list):
    """Por cada {"codigo", "descripcion", "ultimo_precio", "porcentaje",
    "mejor_proveedor"}, hace un upsert en la tabla (crea la fila si no
    existía, o actualiza los datos si ya existía)."""
    if not items:
        return

    hoy = datetime.now(ZONA_HORARIA_ARGENTINA).date()
    filas = [
        (
            item["codigo"],
            item["descripcion"],
            item["ultimo_precio"],
            item.get("porcentaje"),
            item.get("mejor_proveedor"),
            hoy,
        )
        for item in items
    ]

    conn = None
    try:
        conn = _conectar()
        cur = conn.cursor()
        cur.executemany(_UPSERT, filas)
        cur.close()
        conn.commit()
    except ErrorPreciosReferencia:
        raise
    except Exception as exc:
        raise ErrorPreciosReferencia(
            f"No se pudo guardar en la tabla de precios de referencia: {type(exc).__name__}: {exc!r}"
        ) from exc
    finally:
        if conn is not None:
            conn.close()


def guardar_oferta_historial(items: list, encabezado: dict):
    """Guarda la etapa 1 (lo que Zeid ofertó) del historial de
    licitaciones. `items`: lista de {"codigo", "descripcion", "cantidad",
    "precio_ofertado", "porcentaje_ofertado", "proveedor_elegido"}.
    `encabezado`: {"expediente", "contratacion"} del F.41 cargado. No
    hace nada si no hay expediente (no se puede vincular a una
    licitación)."""
    expediente = (encabezado or {}).get("expediente") or ""
    if not items or not expediente:
        return

    contratacion = (encabezado or {}).get("contratacion")
    hoy = datetime.now(ZONA_HORARIA_ARGENTINA).date()
    filas = [
        (
            item["codigo"],
            _normalizar_descripcion(item["descripcion"]),
            expediente,
            item["descripcion"],
            contratacion,
            item.get("cantidad"),
            item.get("precio_ofertado"),
            item.get("porcentaje_ofertado"),
            item.get("proveedor_elegido"),
            hoy,
        )
        for item in items
    ]

    conn = None
    try:
        conn = _conectar()
        cur = conn.cursor()
        cur.executemany(_UPSERT_OFERTA, filas)
        cur.close()
        conn.commit()
    except ErrorPreciosReferencia:
        raise
    except Exception as exc:
        raise ErrorPreciosReferencia(
            f"No se pudo guardar el historial de la oferta: {type(exc).__name__}: {exc!r}"
        ) from exc
    finally:
        if conn is not None:
            conn.close()


def guardar_resultado_adjudicacion(items: list, encabezado: dict):
    """Guarda la etapa 2 (resultado real) y, si vienen, la etapa 3
    (costo real) del historial de licitaciones. `items`: lista de
    {"codigo", "descripcion", "cantidad", "gano_zeid",
    "precio_adjudicado", "proveedor_ganador", "costo_real"?,
    "precio_venta_real"?}. `encabezado`: {"expediente", "contratacion"}
    del F.43 cargado."""
    expediente = (encabezado or {}).get("expediente") or ""
    if not items or not expediente:
        return

    contratacion = (encabezado or {}).get("contratacion")
    hoy = datetime.now(ZONA_HORARIA_ARGENTINA).date()
    filas = [
        (
            item["codigo"],
            _normalizar_descripcion(item["descripcion"]),
            expediente,
            item["descripcion"],
            contratacion,
            item.get("cantidad"),
            item.get("gano_zeid"),
            item.get("precio_adjudicado"),
            item.get("proveedor_ganador"),
            hoy,
            item.get("costo_real"),
            item.get("precio_venta_real"),
            hoy if (item.get("costo_real") is not None or item.get("precio_venta_real") is not None) else None,
        )
        for item in items
    ]

    conn = None
    try:
        conn = _conectar()
        cur = conn.cursor()
        cur.executemany(_UPSERT_ADJUDICACION, filas)
        cur.close()
        conn.commit()
    except ErrorPreciosReferencia:
        raise
    except Exception as exc:
        raise ErrorPreciosReferencia(
            f"No se pudo guardar el resultado de la adjudicación: {type(exc).__name__}: {exc!r}"
        ) from exc
    finally:
        if conn is not None:
            conn.close()


def leer_resumen_proveedores() -> dict:
    """Devuelve {"proveedores": [{"proveedor", "veces_elegido",
    "veces_ganadas", "porcentaje_exito"}, ...] (ordenado por veces
    elegido, descendente), "totales": {"licitaciones_ganadas",
    "costo_total", "venta_total", "margen_total"}}."""
    conn = None
    try:
        conn = _conectar()
        cur = conn.cursor()
        cur.execute(
            "SELECT proveedor_elegido, COUNT(*), COUNT(*) FILTER (WHERE gano_zeid) "
            "FROM historial_licitaciones "
            "WHERE proveedor_elegido IS NOT NULL AND proveedor_elegido <> '' "
            "GROUP BY proveedor_elegido "
            "ORDER BY COUNT(*) DESC;"
        )
        filas_proveedores = cur.fetchall()
        cur.execute(
            "SELECT COUNT(*) FILTER (WHERE gano_zeid), "
            "SUM(costo_real * cantidad) FILTER (WHERE gano_zeid), "
            "SUM(precio_venta_real * cantidad) FILTER (WHERE gano_zeid) "
            "FROM historial_licitaciones;"
        )
        licitaciones_ganadas, costo_total, venta_total = cur.fetchone()
        cur.close()
    except ErrorPreciosReferencia:
        raise
    except Exception as exc:
        raise ErrorPreciosReferencia(
            f"No se pudo leer el resumen de proveedores: {type(exc).__name__}: {exc!r}"
        ) from exc
    finally:
        if conn is not None:
            conn.close()

    proveedores = [
        {
            "proveedor": proveedor,
            "veces_elegido": veces_elegido,
            "veces_ganadas": veces_ganadas,
            "porcentaje_exito": round(100 * veces_ganadas / veces_elegido, 1) if veces_elegido else 0,
        }
        for proveedor, veces_elegido, veces_ganadas in filas_proveedores
    ]

    costo_total = float(costo_total) if costo_total is not None else 0.0
    venta_total = float(venta_total) if venta_total is not None else 0.0
    return {
        "proveedores": proveedores,
        "totales": {
            "licitaciones_ganadas": licitaciones_ganadas or 0,
            "costo_total": costo_total,
            "venta_total": venta_total,
            "margen_total": venta_total - costo_total,
        },
    }


def leer_historial_por_expediente(expediente: str) -> list:
    """Todas las filas de historial_licitaciones para un expediente
    puntual (sin la columna interna descripcion_normalizada), ordenadas
    por código."""
    conn = None
    try:
        conn = _conectar()
        cur = conn.cursor()
        cur.execute(
            "SELECT codigo, descripcion, cantidad, precio_ofertado, proveedor_elegido, "
            "gano_zeid, precio_adjudicado, proveedor_ganador, costo_real, precio_venta_real "
            "FROM historial_licitaciones WHERE expediente = %s ORDER BY codigo;",
            (expediente,),
        )
        filas = cur.fetchall()
        cur.close()
    except ErrorPreciosReferencia:
        raise
    except Exception as exc:
        raise ErrorPreciosReferencia(
            f"No se pudo leer el historial del expediente: {type(exc).__name__}: {exc!r}"
        ) from exc
    finally:
        if conn is not None:
            conn.close()

    return [
        {
            "codigo": codigo,
            "descripcion": descripcion,
            "cantidad": float(cantidad) if cantidad is not None else None,
            "precio_ofertado": float(precio_ofertado) if precio_ofertado is not None else None,
            "proveedor_elegido": proveedor_elegido,
            "gano_zeid": gano_zeid,
            "precio_adjudicado": float(precio_adjudicado) if precio_adjudicado is not None else None,
            "proveedor_ganador": proveedor_ganador,
            "costo_real": float(costo_real) if costo_real is not None else None,
            "precio_venta_real": float(precio_venta_real) if precio_venta_real is not None else None,
        }
        for (
            codigo, descripcion, cantidad, precio_ofertado, proveedor_elegido,
            gano_zeid, precio_adjudicado, proveedor_ganador, costo_real, precio_venta_real,
        ) in filas
    ]


def leer_historial_completo() -> list:
    """Todas las filas de historial_licitaciones, todas las columnas
    visibles (sin la interna descripcion_normalizada), para exportar a
    Excel. Devuelve tuplas crudas (no dicts) en el mismo orden que
    ENCABEZADOS_HISTORIAL, listas para volcar directo a una hoja."""
    conn = None
    try:
        conn = _conectar()
        cur = conn.cursor()
        cur.execute(
            "SELECT codigo, descripcion, expediente, contratacion, cantidad, "
            "precio_ofertado, porcentaje_ofertado, proveedor_elegido, fecha_oferta, "
            "gano_zeid, precio_adjudicado, proveedor_ganador, fecha_adjudicacion, "
            "costo_real, precio_venta_real, fecha_costo "
            "FROM historial_licitaciones ORDER BY expediente, codigo;"
        )
        filas = cur.fetchall()
        cur.close()
    except ErrorPreciosReferencia:
        raise
    except Exception as exc:
        raise ErrorPreciosReferencia(
            f"No se pudo leer el historial completo: {type(exc).__name__}: {exc!r}"
        ) from exc
    finally:
        if conn is not None:
            conn.close()
    return filas


ENCABEZADOS_HISTORIAL = [
    "Codigo", "Descripcion", "Expediente", "Contratacion", "Cantidad",
    "Precio ofertado", "% ofertado", "Proveedor elegido", "Fecha oferta",
    "Gano Zeid", "Precio adjudicado", "Proveedor ganador", "Fecha adjudicacion",
    "Costo real", "Precio venta real", "Fecha costo",
]


def escribir_historial_excel(filas: list, salida: Path):
    """Vuelca a un .xlsx todo el detalle crudo de historial_licitaciones
    (una fila por ítem+licitación), para archivar o analizar aparte.
    `filas` es la lista de tuplas que devuelve leer_historial_completo."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Historial"

    fuente_encabezado = Font(bold=True, color="000000")
    relleno_encabezado = PatternFill("solid", fgColor="D4EA6B")
    centrado = Alignment(horizontal="center")

    for col, titulo in enumerate(ENCABEZADOS_HISTORIAL, start=1):
        c = ws.cell(row=1, column=col, value=titulo)
        c.font = fuente_encabezado
        c.fill = relleno_encabezado
        c.alignment = centrado

    columnas_moneda = (6, 11, 14, 15)  # Precio ofertado, Precio adjudicado, Costo real, Precio venta real
    columnas_fecha = (9, 13, 16)
    columna_booleana = 10  # Gano Zeid
    fila_excel = 2
    for fila in filas:
        for col, valor in enumerate(fila, start=1):
            if col in columnas_fecha and valor is not None:
                valor = valor.isoformat()
            elif col == columna_booleana and valor is not None:
                valor = "Si" if valor else "No"
            c = ws.cell(row=fila_excel, column=col, value=valor)
            c.alignment = Alignment(horizontal="left", wrap_text=True) if col == 2 else centrado
            if col in columnas_moneda and valor is not None:
                c.number_format = FMT_MONEDA_ARS
        fila_excel += 1

    anchos = {
        "A": 14, "B": 46, "C": 14, "D": 14, "E": 10, "F": 14, "G": 10,
        "H": 24, "I": 12, "J": 10, "K": 14, "L": 24, "M": 14, "N": 12, "O": 14, "P": 12,
    }
    for letra, ancho in anchos.items():
        ws.column_dimensions[letra].width = ancho

    wb.save(salida)
