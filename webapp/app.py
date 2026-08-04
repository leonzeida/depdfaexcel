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
    actualizar_precios_referencia,
    escribir_comparacion_precios,
    escribir_notas_pedido,
    escribir_planilla_trabajo,
    expediente_slug,
    extraer_encabezado,
    extraer_pdf,
    leer_precios_referencia,
    nombre_archivo_notas_pedido,
)

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


@app.route("/comparar", methods=["POST"])
def comparar():
    archivo = request.files.get("pdf")
    if not archivo or archivo.filename == "":
        return jsonify(error="Elegí un archivo PDF antes de comparar."), 400

    if not archivo.filename.lower().endswith(".pdf"):
        return jsonify(error="El archivo tiene que ser un PDF."), 400

    maestro = request.files.get("maestro")

    with tempfile.TemporaryDirectory() as tmp:
        pdf_path = Path(tmp) / archivo.filename
        archivo.save(pdf_path)

        maestro_path = None
        if maestro and maestro.filename:
            maestro_path = Path(tmp) / maestro.filename
            maestro.save(maestro_path)

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
            referencias = leer_precios_referencia(maestro_path)
        except Exception:
            return jsonify(
                error="No se pudo leer el archivo de precios de referencia. Revisá que sea el maestro correcto."
            ), 400

        slug = expediente_slug(encabezado["expediente"])
        salida = Path(tmp) / f"Comparacion de precios {slug}.xlsx"
        escribir_comparacion_precios(filas, encabezado, salida, referencias)

        datos = base64.b64encode(salida.read_bytes()).decode("ascii")
        return jsonify(
            archivos=[{"etiqueta": "Comparación de precios", "nombre": salida.name, "datos": datos}]
        )


@app.route("/actualizar_referencia", methods=["POST"])
def actualizar_referencia():
    comparacion = request.files.get("comparacion")
    if not comparacion or comparacion.filename == "":
        return jsonify(error="Elegí el archivo de comparación de precios ya completado."), 400

    maestro = request.files.get("maestro")

    with tempfile.TemporaryDirectory() as tmp:
        comparacion_path = Path(tmp) / comparacion.filename
        comparacion.save(comparacion_path)

        maestro_path = None
        if maestro and maestro.filename:
            maestro_path = Path(tmp) / maestro.filename
            maestro.save(maestro_path)

        salida = Path(tmp) / "Precios de referencia.xlsx"
        try:
            actualizar_precios_referencia(comparacion_path, maestro_path, salida)
        except Exception:
            return jsonify(
                error="No se pudo leer ese archivo. Revisá que sea una comparación de precios generada por esta página."
            ), 400

        datos = base64.b64encode(salida.read_bytes()).decode("ascii")
        return jsonify(
            archivos=[{"etiqueta": "Precios de referencia", "nombre": salida.name, "datos": datos}]
        )


def abrir_navegador():
    webbrowser.open(f"http://127.0.0.1:{PUERTO}")


if __name__ == "__main__":
    threading.Timer(1.0, abrir_navegador).start()
    app.run(host="127.0.0.1", port=PUERTO, debug=False)
