#!/usr/bin/env python3
"""Pagina web local para convertir un F.41 (PDF) a Excel.

Se ejecuta en esta misma PC y se abre en el navegador. No necesita
internet ni instalar nada mas: usa el mismo motor de extraccion que
f41_a_excel.py.
"""

import sys
import tempfile
import threading
import webbrowser
from pathlib import Path

from flask import Flask, jsonify, render_template, request, send_file

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from f41_a_excel import extraer_pdf, escribir_excel  # noqa: E402

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
        except Exception:
            return jsonify(
                error="No se pudo leer ese PDF. Revisá que sea un Pedido de Cotización F.41 válido."
            ), 400

        if not filas:
            return jsonify(error="No se encontraron items en la tabla de ese PDF."), 400

        salida = pdf_path.with_suffix(".xlsx")
        escribir_excel(filas, salida)

        return send_file(
            salida,
            as_attachment=True,
            download_name=salida.name,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


def abrir_navegador():
    webbrowser.open(f"http://127.0.0.1:{PUERTO}")


if __name__ == "__main__":
    threading.Timer(1.0, abrir_navegador).start()
    app.run(host="127.0.0.1", port=PUERTO, debug=False)
