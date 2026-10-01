import json
from datetime import date

from flask import Flask, jsonify, render_template, request, send_file

from procesador import generar_pagos, generar_retiradas


app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 80 * 1024 * 1024


@app.get("/")
def inicio():
    return render_template("index.html", hoy=date.today().isoformat())


def generar_reporte_solicitado():
    faltantes = [
        nombre for nombre in ("nomina", "altas", "base")
        if nombre not in request.files
    ]
    if faltantes:
        raise ValueError("Debes cargar Nómina, Altas y Reporte revisión PBS.")

    tipo = request.form.get("tipo", "pagos")
    fecha_corte = request.form.get("fecha_corte", date.today().isoformat())
    nomina = request.files["nomina"].read()
    altas = request.files["altas"].read()
    base = request.files["base"].read()

    if tipo == "retiradas":
        archivo, datos = generar_retiradas(nomina, base)
        nombre = "retiradas.xlsx"
    elif tipo == "pagos":
        archivo, datos = generar_pagos(nomina, altas, base, fecha_corte)
        nombre = "pendientes_no_pendientes.xlsx"
    else:
        raise ValueError("El tipo de reporte seleccionado no es válido.")

    return archivo, datos, nombre


@app.post("/procesar")
def procesar():
    _, datos, nombre = generar_reporte_solicitado()
    registros = json.loads(datos.to_json(orient="records", date_format="iso"))
    return jsonify(
        columnas=list(datos.columns),
        registros=registros,
        total=len(datos),
        nombre_archivo=nombre,
    )


@app.post("/exportar")
def exportar():
    archivo, datos, nombre = generar_reporte_solicitado()
    respuesta = send_file(
        archivo,
        as_attachment=True,
        download_name=nombre,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    respuesta.headers["X-Registros-Procesados"] = str(len(datos))
    return respuesta


@app.errorhandler(Exception)
def manejar_error(error):
    if app.debug:
        raise error
    return jsonify(error=str(error)), 500


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False)
