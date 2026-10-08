#!/usr/bin/env python3
"""Pagina web local para convertir un F.41 (PDF) a Excel.

Se ejecuta en esta misma PC y se abre en el navegador. No necesita
internet ni instalar nada mas: usa el mismo motor de extraccion que
f41_a_excel.py.
"""

import base64
import hmac
import json
import os
import secrets
import sys
import tempfile
import threading
import webbrowser
from pathlib import Path

from flask import Flask, jsonify, redirect, render_template, request, session, url_for

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from f41_a_excel import (  # noqa: E402
    PREFIJOS_CODIGO,
    clave_item,
    escribir_comparacion_precios_excel,
    escribir_notas_pedido,
    escribir_planilla_trabajo,
    expediente_slug,
    extraer_encabezado,
    extraer_pdf,
    extraer_pdf_completo,
    linea_de_pedido,
    nombre_archivo_notas_pedido,
)
from db_referencia import (
    ErrorPreciosReferencia,
    escribir_historial_excel,
    guardar_oferta_historial,
    guardar_precios_referencia,
    guardar_resultado_adjudicacion,
    leer_historial_completo,
    leer_historial_por_expediente,
    leer_precios_referencia,
    leer_resumen_proveedores,
)
from editar_f41_pdf import generar_f41_editado
from f43_preadjudicacion import (
    es_zeid_medical,
    escribir_preadjudicacion_excel,
    extraer_encabezado_f43,
    extraer_preadjudicacion_completa,
    titulo_preadjudicacion,
)

app = Flask(__name__)
PUERTO = 5000

# Si SECRET_KEY no está seteada en el entorno (ej. en Render), se genera
# una al arrancar el proceso - las sesiones se resetean en cada reinicio
# del servidor, pero eso es aceptable para este caso de uso (si hiciera
# falta que no pase, hay que setear SECRET_KEY en el entorno).
app.secret_key = os.environ.get("SECRET_KEY") or secrets.token_hex(32)

SITE_PASSWORD = os.environ.get("SITE_PASSWORD")
if not SITE_PASSWORD:
    print(
        "[login] AVISO: no está configurada SITE_PASSWORD - el sitio queda "
        "sin contraseña, cualquiera con el link puede entrar.",
        file=sys.stderr,
        flush=True,
    )

# Rutas que no requieren haber iniciado sesión.
_RUTAS_PUBLICAS = {"/login", "/static"}


@app.before_request
def exigir_login():
    if not SITE_PASSWORD:
        return  # sin contraseña configurada, no se bloquea nada
    if request.path in _RUTAS_PUBLICAS or request.path.startswith("/static/"):
        return
    if not session.get("autenticado"):
        return redirect(url_for("login"))


@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        intento = request.form.get("password", "")
        if SITE_PASSWORD and hmac.compare_digest(intento, SITE_PASSWORD):
            session["autenticado"] = True
            return redirect(url_for("inicio"))
        error = "Contraseña incorrecta."
    return render_template("login.html", error=error)


@app.route("/logout", methods=["GET"])
def logout():
    session.clear()
    return redirect(url_for("login"))


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
            todas_las_filas = extraer_pdf_completo(pdf_path)
            encabezado = extraer_encabezado(pdf_path)
        except Exception:
            return jsonify(
                error="No se pudo leer ese PDF. Revisá que sea un Pedido de Cotización F.41 válido."
            ), 400

        filas = [f for f in todas_las_filas if f["codigo"].startswith(PREFIJOS_CODIGO)]
        otras_filas = [f for f in todas_las_filas if not f["codigo"].startswith(PREFIJOS_CODIGO)]

        if not filas:
            return jsonify(error="No se encontraron items en la tabla de ese PDF."), 400

    try:
        referencias = leer_precios_referencia()
    except ErrorPreciosReferencia as exc:
        print(f"[precios-referencia] Fallo leyendo: {exc}", file=sys.stderr, flush=True)
        return jsonify(
            error="No se pudo conectar con la base de precios de referencia. Probá de nuevo en un momento."
        ), 502

    def _mapear_item(item):
        datos = referencias.get(clave_item(item["codigo"], item["descripcion"]), {})
        return {
            "rg": item["rg"],
            "codigo": item["codigo"],
            "descripcion": item["descripcion"],
            "cantidad": item["cantidad"],
            "ultimo_precio": datos.get("ultimo_precio"),
            "porcentaje": datos.get("porcentaje"),
            "actualizado": datos.get("actualizado"),
            "mejor_proveedor": datos.get("mejor_proveedor"),
        }

    items = [_mapear_item(item) for item in filas]
    # Ítems del mismo PDF cuyo código no es de las categorías habituales
    # (PREFIJOS_CODIGO): el comparador los deja afuera de la carga
    # automática, pero el usuario los puede buscar y agregar a mano con
    # el botón "Agregar ítem".
    otros_items = [_mapear_item(item) for item in otras_filas]

    return jsonify(
        titulo=linea_de_pedido(encabezado, incluir_titulo=False),
        items=items,
        otros_items=otros_items,
        encabezado=encabezado,
    )


@app.route("/guardar_referencia", methods=["POST"])
def guardar_referencia():
    cuerpo = request.get_json(silent=True) or {}
    items = cuerpo.get("items") or []
    encabezado = cuerpo.get("encabezado") or {}

    items_validos = []
    for item in items:
        codigo = (item.get("codigo") or "").strip()
        descripcion = (item.get("descripcion") or "").strip()
        ultimo_precio = item.get("ultimo_precio")
        if not codigo or ultimo_precio is None:
            continue
        try:
            ultimo_precio = float(ultimo_precio)
        except (TypeError, ValueError):
            continue

        porcentaje = item.get("porcentaje")
        if porcentaje is not None:
            try:
                porcentaje = float(porcentaje)
            except (TypeError, ValueError):
                porcentaje = None

        mejor_proveedor = (item.get("mejor_proveedor") or "").strip() or None

        cantidad = item.get("cantidad")
        try:
            cantidad = float(cantidad) if cantidad is not None else None
        except (TypeError, ValueError):
            cantidad = None

        items_validos.append(
            {
                "codigo": codigo,
                "descripcion": descripcion,
                "ultimo_precio": ultimo_precio,
                "porcentaje": porcentaje,
                "mejor_proveedor": mejor_proveedor,
                "cantidad": cantidad,
            }
        )

    if not items_validos:
        return jsonify(error="No hay ningún precio para guardar."), 400

    try:
        guardar_precios_referencia(items_validos)
        # El historial de licitaciones es una ampliación aparte de
        # precios_referencia (no la reemplaza): si falla no se corta el
        # guardado principal, que es el que ya usaba el comparador.
        try:
            items_oferta = [
                {
                    "codigo": it["codigo"],
                    "descripcion": it["descripcion"],
                    "cantidad": it["cantidad"],
                    "precio_ofertado": it["ultimo_precio"],
                    "porcentaje_ofertado": it["porcentaje"],
                    "proveedor_elegido": it["mejor_proveedor"],
                }
                for it in items_validos
            ]
            guardar_oferta_historial(items_oferta, encabezado)
        except ErrorPreciosReferencia as exc:
            print(f"[historial-licitaciones] Fallo guardando oferta: {exc}", file=sys.stderr, flush=True)
    except ErrorPreciosReferencia as exc:
        print(f"[precios-referencia] Fallo guardando: {exc}", file=sys.stderr, flush=True)
        return jsonify(
            error="No se pudo guardar en la base de precios de referencia. Probá de nuevo en un momento."
        ), 502

    return jsonify(ok=True, cantidad=len(items_validos))


@app.route("/exportar_comparacion", methods=["POST"])
def exportar_comparacion():
    cuerpo = request.get_json(silent=True) or {}
    titulo = (cuerpo.get("titulo") or "").strip()
    filas = cuerpo.get("filas") or []

    if not isinstance(filas, list) or not filas:
        return jsonify(error="No hay datos en la grilla para exportar."), 400

    with tempfile.TemporaryDirectory() as tmp:
        salida = Path(tmp) / "Comparacion de precios.xlsx"
        try:
            escribir_comparacion_precios_excel(titulo, filas, salida)
        except Exception as exc:
            print(f"[exportar-comparacion] Fallo generando el Excel: {exc}", file=sys.stderr, flush=True)
            return jsonify(error="No se pudo generar el Excel de la comparación."), 500

        datos = base64.b64encode(salida.read_bytes()).decode("ascii")

    return jsonify(archivos=[{"etiqueta": "Comparación", "nombre": salida.name, "datos": datos}])


@app.route("/items_para_editor", methods=["POST"])
def items_para_editor():
    archivo = request.files.get("pdf")
    if not archivo or archivo.filename == "":
        return jsonify(error="Elegí un archivo PDF antes de continuar."), 400

    if not archivo.filename.lower().endswith(".pdf"):
        return jsonify(error="El archivo tiene que ser un PDF."), 400

    with tempfile.TemporaryDirectory() as tmp:
        pdf_path = Path(tmp) / archivo.filename
        archivo.save(pdf_path)

        try:
            todas_las_filas = extraer_pdf_completo(pdf_path)
        except Exception:
            return jsonify(
                error="No se pudo leer ese PDF. Revisá que sea un Pedido de Cotización F.41 válido."
            ), 400

        if not todas_las_filas:
            return jsonify(error="No se encontraron items en la tabla de ese PDF."), 400

    items = [
        {
            "rg": item["rg"],
            "codigo": item["codigo"],
            "descripcion": item["descripcion"],
            "cantidad": item["cantidad"],
        }
        for item in todas_las_filas
    ]

    return jsonify(items=items)


@app.route("/generar_f41_editado", methods=["POST"])
def generar_f41_editado_ruta():
    archivo = request.files.get("pdf")
    if not archivo or archivo.filename == "":
        return jsonify(error="Hace falta volver a adjuntar el PDF original para generar el F.41 editado."), 400

    if not archivo.filename.lower().endswith(".pdf"):
        return jsonify(error="El archivo tiene que ser un PDF."), 400

    try:
        filas = json.loads(request.form.get("filas", ""))
    except (TypeError, ValueError):
        filas = None

    if not isinstance(filas, list) or not filas:
        return jsonify(error="No hay renglones cargados para generar el PDF."), 400

    with tempfile.TemporaryDirectory() as tmp:
        pdf_path = Path(tmp) / archivo.filename
        archivo.save(pdf_path)

        salida = Path(tmp) / "F41 editado.pdf"
        try:
            generar_f41_editado(pdf_path, filas, salida)
        except Exception as exc:
            print(f"[generar-f41-editado] Fallo generando el PDF: {exc}", file=sys.stderr, flush=True)
            return jsonify(
                error="No se pudo generar el F.41 editado. Revisá que el PDF original sea un F.41 válido."
            ), 500

        datos = base64.b64encode(salida.read_bytes()).decode("ascii")

    return jsonify(archivos=[{"etiqueta": "F.41 editado", "nombre": salida.name, "datos": datos}])


@app.route("/items_para_preadjudicacion", methods=["POST"])
def items_para_preadjudicacion():
    archivo = request.files.get("pdf")
    if not archivo or archivo.filename == "":
        return jsonify(error="Elegí un archivo PDF antes de continuar."), 400

    if not archivo.filename.lower().endswith(".pdf"):
        return jsonify(error="El archivo tiene que ser un PDF."), 400

    with tempfile.TemporaryDirectory() as tmp:
        pdf_path = Path(tmp) / archivo.filename
        archivo.save(pdf_path)

        try:
            items = extraer_preadjudicacion_completa(pdf_path)
            encabezado = extraer_encabezado_f43(pdf_path)
        except Exception:
            return jsonify(
                error="No se pudo leer ese PDF. Revisá que sea un Acta de Preadjudicación F.43 válida."
            ), 400

        if not items:
            return jsonify(error="No se encontraron items en ese PDF."), 400

    return jsonify(titulo=titulo_preadjudicacion(encabezado), items=items)


@app.route("/exportar_preadjudicacion", methods=["POST"])
def exportar_preadjudicacion():
    cuerpo = request.get_json(silent=True) or {}
    titulo = (cuerpo.get("titulo") or "").strip()
    filas = cuerpo.get("filas") or []

    if not isinstance(filas, list) or not filas:
        return jsonify(error="No hay datos en la grilla para exportar."), 400

    with tempfile.TemporaryDirectory() as tmp:
        salida = Path(tmp) / "Precios adjudicados.xlsx"
        try:
            escribir_preadjudicacion_excel(titulo, filas, salida)
        except Exception as exc:
            print(f"[exportar-preadjudicacion] Fallo generando el Excel: {exc}", file=sys.stderr, flush=True)
            return jsonify(error="No se pudo generar el Excel de precios adjudicados."), 500

        datos = base64.b64encode(salida.read_bytes()).decode("ascii")

    return jsonify(archivos=[{"etiqueta": "Precios adjudicados", "nombre": salida.name, "datos": datos}])


@app.route("/generar_nota_pedido_preadjudicacion", methods=["POST"])
def generar_nota_pedido_preadjudicacion():
    archivo = request.files.get("pdf")
    if not archivo or archivo.filename == "":
        return jsonify(error="Hace falta volver a adjuntar el PDF del acta para generar la Nota de Pedido."), 400

    if not archivo.filename.lower().endswith(".pdf"):
        return jsonify(error="El archivo tiene que ser un PDF."), 400

    with tempfile.TemporaryDirectory() as tmp:
        pdf_path = Path(tmp) / archivo.filename
        archivo.save(pdf_path)

        try:
            items = extraer_preadjudicacion_completa(pdf_path)
            encabezado = extraer_encabezado_f43(pdf_path)
        except Exception:
            return jsonify(
                error="No se pudo leer ese PDF. Revisá que sea un Acta de Preadjudicación F.43 válida."
            ), 400

        items_zeid = [it for it in items if es_zeid_medical(it["proveedor"])]
        if not items_zeid:
            return jsonify(error="Zeid Medical no ganó ningún renglón en esta acta."), 400

        filas = [
            {"rg": it["renglon"], "descripcion": it["descripcion"], "cantidad": it["cantidad"]}
            for it in items_zeid
        ]

        try:
            salida = Path(tmp) / f"{nombre_archivo_notas_pedido(encabezado)}.xlsx"
            escribir_notas_pedido(filas, encabezado, salida)
        except Exception as exc:
            print(f"[nota-pedido-preadjudicacion] Fallo generando el Excel: {exc}", file=sys.stderr, flush=True)
            return jsonify(error="No se pudo generar la Nota de Pedido."), 500

        datos = base64.b64encode(salida.read_bytes()).decode("ascii")

    return jsonify(archivos=[{"etiqueta": "Nota de Pedido", "nombre": salida.name, "datos": datos}])


@app.route("/guardar_resultado_preadjudicacion", methods=["POST"])
def guardar_resultado_preadjudicacion():
    archivo = request.files.get("pdf")
    if not archivo or archivo.filename == "":
        return jsonify(error="Hace falta volver a adjuntar el PDF del acta para guardar el resultado."), 400

    if not archivo.filename.lower().endswith(".pdf"):
        return jsonify(error="El archivo tiene que ser un PDF."), 400

    try:
        costos = json.loads(request.form.get("costos", "") or "{}")
    except (TypeError, ValueError):
        costos = {}

    with tempfile.TemporaryDirectory() as tmp:
        pdf_path = Path(tmp) / archivo.filename
        archivo.save(pdf_path)

        try:
            items = extraer_preadjudicacion_completa(pdf_path)
            encabezado = extraer_encabezado_f43(pdf_path)
        except Exception:
            return jsonify(
                error="No se pudo leer ese PDF. Revisá que sea un Acta de Preadjudicación F.43 válida."
            ), 400

    # Los renglones "DESIERTO" no tienen precio real, no aportan nada al
    # historial.
    items_adjudicados = [it for it in items if it["proveedor"] != "DESIERTO"]
    if not items_adjudicados:
        return jsonify(error="No hay renglones adjudicados en esta acta para guardar."), 400

    def _num(valor):
        try:
            return float(valor) if valor not in (None, "") else None
        except (TypeError, ValueError):
            return None

    items_resultado = []
    for it in items_adjudicados:
        datos_costo = costos.get(str(it["renglon"])) or {}
        items_resultado.append(
            {
                "codigo": it["codigo"],
                "descripcion": it["descripcion"],
                "cantidad": it["cantidad"] if isinstance(it["cantidad"], (int, float)) else None,
                "gano_zeid": es_zeid_medical(it["proveedor"]),
                "precio_adjudicado": it["precio_unitario"],
                "proveedor_ganador": it["proveedor"],
                "costo_real": _num(datos_costo.get("costo")),
                "precio_venta_real": _num(datos_costo.get("precio_venta")),
            }
        )

    try:
        guardar_resultado_adjudicacion(items_resultado, encabezado)
    except ErrorPreciosReferencia as exc:
        print(f"[historial-licitaciones] Fallo guardando resultado: {exc}", file=sys.stderr, flush=True)
        return jsonify(
            error="No se pudo guardar en la base de historial. Probá de nuevo en un momento."
        ), 502

    return jsonify(ok=True, cantidad=len(items_resultado))


@app.route("/historial_resumen", methods=["GET"])
def historial_resumen():
    try:
        resumen = leer_resumen_proveedores()
    except ErrorPreciosReferencia as exc:
        print(f"[historial-licitaciones] Fallo leyendo resumen: {exc}", file=sys.stderr, flush=True)
        return jsonify(
            error="No se pudo conectar con la base de historial. Probá de nuevo en un momento."
        ), 502

    return jsonify(resumen)


@app.route("/historial_expediente", methods=["GET"])
def historial_expediente():
    expediente = (request.args.get("expediente") or "").strip()
    if not expediente:
        return jsonify(error="Escribí un número de expediente para buscar."), 400

    try:
        items = leer_historial_por_expediente(expediente)
    except ErrorPreciosReferencia as exc:
        print(f"[historial-licitaciones] Fallo leyendo expediente: {exc}", file=sys.stderr, flush=True)
        return jsonify(
            error="No se pudo conectar con la base de historial. Probá de nuevo en un momento."
        ), 502

    return jsonify(items=items)


@app.route("/exportar_historial", methods=["GET"])
def exportar_historial():
    try:
        filas = leer_historial_completo()
    except ErrorPreciosReferencia as exc:
        print(f"[historial-licitaciones] Fallo leyendo historial completo: {exc}", file=sys.stderr, flush=True)
        return jsonify(
            error="No se pudo conectar con la base de historial. Probá de nuevo en un momento."
        ), 502

    if not filas:
        return jsonify(error="Todavía no hay nada guardado en el historial para exportar."), 400

    with tempfile.TemporaryDirectory() as tmp:
        salida = Path(tmp) / "Historial de licitaciones.xlsx"
        try:
            escribir_historial_excel(filas, salida)
        except Exception as exc:
            print(f"[historial-licitaciones] Fallo generando el Excel: {exc}", file=sys.stderr, flush=True)
            return jsonify(error="No se pudo generar el Excel del historial."), 500

        datos = base64.b64encode(salida.read_bytes()).decode("ascii")

    return jsonify(archivos=[{"etiqueta": "Historial", "nombre": salida.name, "datos": datos}])


def abrir_navegador():
    webbrowser.open(f"http://127.0.0.1:{PUERTO}")


if __name__ == "__main__":
    threading.Timer(1.0, abrir_navegador).start()
    app.run(host="127.0.0.1", port=PUERTO, debug=False)
