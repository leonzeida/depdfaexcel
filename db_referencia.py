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
    precio_referencia NUMERIC(12,2),
    ultimo_precio NUMERIC(12,2),
    actualizado DATE,
    PRIMARY KEY (codigo, descripcion)
);
"""

_UPSERT = """
INSERT INTO precios_referencia (codigo, descripcion, precio_referencia, ultimo_precio, actualizado)
VALUES (%s, %s, %s, %s, %s)
ON CONFLICT (codigo, descripcion) DO UPDATE SET
    ultimo_precio = EXCLUDED.ultimo_precio,
    precio_referencia = COALESCE(EXCLUDED.precio_referencia, precios_referencia.precio_referencia),
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
    cur.close()
    conn.commit()
    return conn


def leer_precios_referencia() -> dict:
    """Devuelve {(codigo, descripcion): {"precio_referencia": float|None, "ultimo_precio": float|None}}."""
    conn = None
    try:
        conn = _conectar()
        cur = conn.cursor()
        cur.execute("SELECT codigo, descripcion, precio_referencia, ultimo_precio FROM precios_referencia;")
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
    for codigo, descripcion, precio_referencia, ultimo_precio in filas:
        precios[clave_item(codigo, descripcion)] = {
            "precio_referencia": float(precio_referencia) if precio_referencia is not None else None,
            "ultimo_precio": float(ultimo_precio) if ultimo_precio is not None else None,
        }
    return precios


def guardar_precios_referencia(items: list):
    """Por cada {"codigo", "descripcion", "precio_referencia", "ultimo_precio"},
    hace un upsert en la tabla. precio_referencia puede venir en None (no se
    pisa el valor ya guardado); ultimo_precio siempre se guarda tal cual."""
    if not items:
        return

    hoy = datetime.now(ZONA_HORARIA_ARGENTINA).date()
    filas = [
        (item["codigo"], item["descripcion"], item.get("precio_referencia"), item["ultimo_precio"], hoy)
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
