from flask import Flask, request, render_template_string, send_file, redirect, url_for
import pandas as pd
import os
import uuid
import traceback
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet

app = Flask(__name__)

UPLOAD_FOLDER = "temp"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

HTML = """
<!DOCTYPE html>
<html>
<head>
<style>
body { font-family: Arial; background: #f2f2f2; text-align: center; }
.box { background: white; padding: 20px; margin: 40px auto; width: 500px; border-radius: 10px; }
button { padding: 10px; margin: 5px; width: 80%; border-radius: 8px; border: none; background: green; color: white; }
.error { color: red; }
</style>
</head>
<body>

<div class="box">
<h2>📊 Subir Excel</h2>

{% if not alumnos %}
<form method="POST" enctype="multipart/form-data" action="/">
<input type="file" name="file" required>
<button type="submit">Procesar</button>
</form>
{% endif %}

</div>

{% if error %}
<div class="box error">{{error}}</div>
{% endif %}

{% if alumnos %}
<div class="box">
<h3>Selecciona alumno</h3>
{% for alumno in alumnos %}
<form action="/reporte" method="POST">
<input type="hidden" name="alumno" value="{{alumno}}">
<input type="hidden" name="file_id" value="{{file_id}}">
<button>{{alumno}}</button>
</form>
{% endfor %}
</div>
{% endif %}

</body>
</html>
"""

# ✅ PROCESAR EXCEL (INTELIGENTE)
def procesar_excel(path):
    df = pd.read_excel(path, header=None)

    cursos = df.iloc[0].ffill()
    practicas = df.iloc[1]

    # 🔥 detectar columna alumno automáticamente
    col_alumno = None
    for col in df.columns:
        valores = df[col].astype(str)
        if valores.str.contains("ALUMNO|NOMBRE", case=False).any():
            col_alumno = col
            break

    if col_alumno is None:
        col_alumno = 0

    columnas = []

    for i in range(len(cursos)):
        curso = str(cursos[i]).strip()
        practica = str(practicas[i]).strip().upper().replace(" ", "")

        if i == col_alumno:
            columnas.append("Alumno")
        elif practica.startswith("P"):
            columnas.append(f"{curso}_{practica}")
        else:
            columnas.append(None)

    df.columns = columnas

    df = df.iloc[2:]
    df = df.loc[:, df.columns.notna()]
    df = df[df["Alumno"].notna()]

    return df


# 📄 GENERAR PDF
def generar_pdf(alumno, df, file_id):
    data = df[df["Alumno"] == alumno]

    if data.empty:
        raise Exception("Alumno sin datos")

    file_path = f"{UPLOAD_FOLDER}/{file_id}_{alumno}.pdf"

    doc = SimpleDocTemplate(file_path, pagesize=letter)
    styles = getSampleStyleSheet()
    elements = []

    elements.append(Paragraph(f"Alumno: {alumno}", styles["Title"]))
    elements.append(Spacer(1, 20))

    cursos_dict = {}

    for col in df.columns:
        if col != "Alumno":
            curso, _ = col.split("_")
            cursos_dict.setdefault(curso, []).append(col)

    for curso, cols in cursos_dict.items():

        cols = sorted(cols, key=lambda x: int(x.split("_")[1].replace("P","")))

        tabla = [["Práctica", "Nota"]]

        for col in cols:
            val = data.iloc[0][col]
            tabla.append([col.split("_")[1], val])

        t = Table(tabla)
        t.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 1, colors.black)
        ]))

        elements.append(Paragraph(curso, styles["Heading2"]))
        elements.append(t)
        elements.append(Spacer(1, 10))

    doc.build(elements)
    return file_path


# 🌐 HOME
@app.route("/", methods=["GET", "POST"])
def index():
    try:
        if request.method == "POST":

            if "file" not in request.files:
                return render_template_string(HTML, error="No se envió archivo")

            file = request.files["file"]

            if file.filename == "":
                return render_template_string(HTML, error="Selecciona un archivo")

            file_id = str(uuid.uuid4())
            path = os.path.join(UPLOAD_FOLDER, f"{file_id}.xlsx")

            file.save(path)

            if not os.path.exists(path):
                return render_template_string(HTML, error="Error guardando archivo")

            return redirect(url_for("ver_alumnos", file_id=file_id))

        return render_template_string(HTML)

    except Exception as e:
        return render_template_string(HTML, error=str(e))


# 📋 VER ALUMNOS
@app.route("/alumnos/<file_id>")
def ver_alumnos(file_id):
    try:
        path = os.path.join(UPLOAD_FOLDER, f"{file_id}.xlsx")

        if not os.path.exists(path):
            return render_template_string(HTML, error="Archivo no encontrado")

        df = procesar_excel(path)

        alumnos = df["Alumno"].dropna().unique().tolist()

        if not alumnos:
            return render_template_string(HTML, error="No se encontraron alumnos")

        return render_template_string(HTML, alumnos=alumnos, file_id=file_id)

    except Exception as e:
        return render_template_string(HTML, error=str(e))


# 📄 REPORTE
@app.route("/reporte", methods=["POST"])
def reporte():
    try:
        alumno = request.form["alumno"]
        file_id = request.form["file_id"]

        path = os.path.join(UPLOAD_FOLDER, f"{file_id}.xlsx")

        if not os.path.exists(path):
            return "El archivo ya no existe. Súbelo otra vez."

        df = procesar_excel(path)

        pdf = generar_pdf(alumno, df, file_id)

        return send_file(pdf, as_attachment=True)

    except Exception as e:
        return f"Error: {str(e)}<br><pre>{traceback.format_exc()}</pre>"


if __name__ == "__main__":
    app.run(debug=True, use_reloader=False)
