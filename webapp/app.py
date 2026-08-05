#!/usr/bin/env python3
"""Pagina web local para convertir un F.41 (PDF) a Excel.

Se ejecuta en esta misma PC y se abre en el navegador. No necesita
internet ni instalar nada mas: usa el mismo motor de extraccion que
f41_a_excel.py.
"""

import base64
import sys
import tempfile
import threading
import webbrowser
from pathlib import Path

from flask import Flask, jsonify, render_template, request

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from f41_a_excel import (  # noqa: E402
    clave_item,
    escribir_notas_pedido,
    escribir_planilla_trabajo,
    expediente_slug,
    extraer_encabezado,
    extraer_pdf,
    linea_de_pedido,
    nombre_archivo_notas_pedido,
)
from sheets_referencia import ErrorPreciosReferencia, guardar_precios_referencia, leer_precios_referencia

app = Flask(__name__)
PUERTO = 5000


@app.route("/", methods=["GET"])
def inicio():
    return render_template("index.html")


@app.route("/convertir", methods=["POST"])
def convertir():
    archivo = request.files.get("pdf")
    if not archivo or archivo.filename == "":
        return jsonify(error="Elegí un archivo PDF antes de convertir."), 400

    if not archivo.filename.lower().endswith(".pdf"):
        return jsonify(error="El archivo tiene que ser un PDF."), 400

    with tempfile.TemporaryDirectory() as tmp:
        pdf_path = Path(tmp) / archivo.filename
        archivo.save(pdf_path)

        try:
            filas = extraer_pdf(pdf_path)
            encabezado = extraer_encabezado(pdf_path)
        except Exception:
            return jsonify(
                error="No se pudo leer ese PDF. Revisá que sea un Pedido de Cotización F.41 válido."
            ), 400

        if not filas:
            return jsonify(error="No se encontraron items en la tabla de ese PDF."), 400

        slug = expediente_slug(encabezado["expediente"])
        salida_np = Path(tmp) / f"{nombre_archivo_notas_pedido(encabezado)}.xlsx"
        salida_pt = Path(tmp) / f"Planilla de Trabajo {slug}.xlsx"
        escribir_notas_pedido(filas, encabezado, salida_np)
        escribir_planilla_trabajo(filas, encabezado, salida_pt)

        archivos = []
        for etiqueta, ruta in (
            ("Nota de Pedido", salida_np),
            ("Planilla de Trabajo", salida_pt),
        ):
            datos = base64.b64encode(ruta.read_bytes()).decode("ascii")
            archivos.append({"etiqueta": etiqueta, "nombre": ruta.name, "datos": datos})

        return jsonify(archivos=archivos)


@app.route("/items_para_comparar", methods=["POST"])
def items_para_comparar():
    archivo = request.files.get("pdf")
    if not archivo or archivo.filename == "":
        return jsonify(error="Elegí un archivo PDF antes de comparar."), 400

    if not archivo.filename.lower().endswith(".pdf"):
        return jsonify(error="El archivo tiene que ser un PDF."), 400

    with tempfile.TemporaryDirectory() as tmp:
        pdf_path = Path(tmp) / archivo.filename
        archivo.save(pdf_path)

        try:
            filas = extraer_pdf(pdf_path)
            encabezado = extraer_encabezado(pdf_path)
        except Exception:
            return jsonify(
                error="No se pudo leer ese PDF. Revisá que sea un Pedido de Cotización F.41 válido."
            ), 400

        if not filas:
            return jsonify(error="No se encontraron items en la tabla de ese PDF."), 400

    try:
        referencias = leer_precios_referencia()
    except ErrorPreciosReferencia as exc:
        app.logger.error("Fallo leyendo precios de referencia: %s", exc)
        return jsonify(
            error="No se pudo conectar con la planilla de precios de referencia. Probá de nuevo en un momento."
        ), 502

    items = []
    for item in filas:
        precio_referencia = referencias.get(clave_item(item["codigo"], item["descripcion"]))
        items.append(
            {
                "rg": item["rg"],
                "codigo": item["codigo"],
                "descripcion": item["descripcion"],
                "cantidad": item["cantidad"],
                "precio_referencia": precio_referencia,
            }
        )

    return jsonify(titulo=linea_de_pedido(encabezado, incluir_titulo=False), items=items)


@app.route("/guardar_referencia", methods=["POST"])
def guardar_referencia():
    cuerpo = request.get_json(silent=True) or {}
    items = cuerpo.get("items") or []

    items_validos = []
    for item in items:
        codigo = (item.get("codigo") or "").strip()
        descripcion = (item.get("descripcion") or "").strip()
        precio = item.get("precio")
        if not codigo or precio is None:
            continue
        try:
            precio = float(precio)
        except (TypeError, ValueError):
            continue
        items_validos.append({"codigo": codigo, "descripcion": descripcion, "precio": precio})

    if not items_validos:
        return jsonify(error="No hay ningún precio para guardar."), 400

    try:
        guardar_precios_referencia(items_validos)
    except ErrorPreciosReferencia as exc:
        app.logger.error("Fallo guardando precios de referencia: %s", exc)
        return jsonify(
            error="No se pudo guardar en la planilla de precios de referencia. Probá de nuevo en un momento."
        ), 502

    return jsonify(ok=True, cantidad=len(items_validos))


def abrir_navegador():
    webbrowser.open(f"http://127.0.0.1:{PUERTO}")


if __name__ == "__main__":
    threading.Timer(1.0, abrir_navegador).start()
    app.run(host="127.0.0.1", port=PUERTO, debug=False)
