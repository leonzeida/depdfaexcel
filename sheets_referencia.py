#!/usr/bin/env python3
"""Acceso a la Google Sheet que guarda los precios de referencia del
comparador de precios de proveedores.

Se mantiene separado de f41_a_excel.py a propósito: ese módulo no debería
necesitar credenciales de Google para poder usarse por línea de comandos.

Requiere dos variables de entorno:
- GOOGLE_SERVICE_ACCOUNT_JSON: el contenido completo (como texto) de la
  clave JSON de la cuenta de servicio de Google Cloud.
- GOOGLE_SHEET_ID: el ID de la Google Sheet "Precios de referencia"
  (el valor entre /d/ y /edit en su URL), ya compartida con el mail de esa
  cuenta de servicio con permiso de Editor.
"""

import json
import os
from datetime import datetime
from zoneinfo import ZoneInfo

import gspread
from google.oauth2.service_account import Credentials

from f41_a_excel import clave_item

NOMBRE_HOJA = "Precios de referencia"
ENCABEZADOS = ["Codigo", "Descripcion", "Precio de referencia", "Actualizado"]
ZONA_HORARIA_ARGENTINA = ZoneInfo("America/Argentina/Buenos_Aires")

_ALCANCES = ["https://www.googleapis.com/auth/spreadsheets"]


class ErrorPreciosReferencia(Exception):
    """Cualquier problema al leer o guardar en la Google Sheet."""


def _cliente() -> gspread.Client:
    credenciales_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if not credenciales_json:
        raise ErrorPreciosReferencia("Falta configurar GOOGLE_SERVICE_ACCOUNT_JSON.")
    try:
        info = json.loads(credenciales_json)
        credenciales = Credentials.from_service_account_info(info, scopes=_ALCANCES)
        return gspread.authorize(credenciales)
    except Exception as exc:
        raise ErrorPreciosReferencia(f"Credenciales de Google inválidas: {exc}") from exc


def _hoja():
    sheet_id = os.environ.get("GOOGLE_SHEET_ID")
    if not sheet_id:
        raise ErrorPreciosReferencia("Falta configurar GOOGLE_SHEET_ID.")
    libro = _cliente().open_by_key(sheet_id)
    try:
        return libro.worksheet(NOMBRE_HOJA)
    except gspread.WorksheetNotFound:
        hoja = libro.add_worksheet(title=NOMBRE_HOJA, rows=1, cols=len(ENCABEZADOS))
        hoja.append_row(ENCABEZADOS)
        return hoja


def leer_precios_referencia() -> dict:
    """Devuelve {(codigo, descripcion): precio} con todo lo guardado en la Sheet."""
    try:
        filas = _hoja().get_all_values()
    except ErrorPreciosReferencia:
        raise
    except Exception as exc:
        raise ErrorPreciosReferencia(f"No se pudo leer la planilla de precios de referencia: {exc}") from exc

    precios = {}
    for fila in filas[1:]:  # fila[0] son los encabezados
        codigo = fila[0] if len(fila) > 0 else ""
        descripcion = fila[1] if len(fila) > 1 else ""
        precio_texto = fila[2] if len(fila) > 2 else ""
        if not codigo or not precio_texto:
            continue
        try:
            precio = float(precio_texto)
        except ValueError:
            continue
        precios[clave_item(codigo, descripcion)] = precio
    return precios


def guardar_precios_referencia(items: list):
    """Actualiza o agrega, para cada {"codigo", "descripcion", "precio"},
    una fila en la Sheet. No toca las filas de ítems que no vinieron en
    `items`."""
    if not items:
        return

    try:
        hoja = _hoja()
        filas = hoja.get_all_values()
    except ErrorPreciosReferencia:
        raise
    except Exception as exc:
        raise ErrorPreciosReferencia(f"No se pudo abrir la planilla de precios de referencia: {exc}") from exc

    fila_por_clave = {}
    for indice, fila in enumerate(filas[1:], start=2):  # las filas de Sheets arrancan en 1, y la 1 es encabezado
        codigo = fila[0] if len(fila) > 0 else ""
        descripcion = fila[1] if len(fila) > 1 else ""
        if codigo:
            fila_por_clave[clave_item(codigo, descripcion)] = indice

    hoy = datetime.now(ZONA_HORARIA_ARGENTINA).strftime("%d/%m/%y")
    actualizaciones = []
    filas_nuevas = []

    for item in items:
        codigo = item["codigo"]
        descripcion = item["descripcion"]
        precio = item["precio"]
        clave = clave_item(codigo, descripcion)
        fila_existente = fila_por_clave.get(clave)
        if fila_existente:
            actualizaciones.append(
                {"range": f"A{fila_existente}:D{fila_existente}", "values": [[codigo, descripcion, precio, hoy]]}
            )
        else:
            filas_nuevas.append([codigo, descripcion, precio, hoy])

    try:
        if actualizaciones:
            hoja.batch_update(actualizaciones)
        if filas_nuevas:
            hoja.append_rows(filas_nuevas)
    except Exception as exc:
        raise ErrorPreciosReferencia(f"No se pudo guardar en la planilla de precios de referencia: {exc}") from exc
