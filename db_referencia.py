#!/usr/bin/env python3
"""Acceso a la base de datos que guarda los precios de referencia del
comparador de precios de proveedores.

Se mantiene separado de f41_a_excel.py a propósito: ese módulo no debería
necesitar una base de datos para poder usarse por línea de comandos.

Requiere una variable de entorno:
- DATABASE_URL: connection string de Postgres (ej. la que da Neon), con
  el formato postgres://usuario:contraseña@host/db?sslmode=require.
"""

import os
from datetime import datetime
from zoneinfo import ZoneInfo

import psycopg

from f41_a_excel import clave_item

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
    cur.close()
    conn.commit()
    return conn


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
