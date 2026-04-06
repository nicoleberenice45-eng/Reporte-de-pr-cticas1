from flask import Flask, request, render_template_string, send_file, redirect, url_for
import pandas as pd
import os
import uuid
import traceback

from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    Frame, PageTemplate, KeepTogether, PageBreak, NextPageTemplate
)
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

# 📊 PROCESAR EXCEL
def procesar_excel(path):
    df = pd.read_excel(path, header=None)
    df = df.dropna(how="all").reset_index(drop=True)

    fila_practicas = None
    for i in range(len(df)):
        fila = df.iloc[i].astype(str).str.upper()
        if fila.str.contains("P1").any():
            fila_practicas = i
            break

    if fila_practicas is None:
        raise Exception("No se encontró fila de prácticas")

    fila_cursos = fila_practicas - 1

    cursos = df.iloc[fila_cursos].ffill()
    practicas = df.iloc[fila_practicas]

    columnas = []

    for i in range(len(cursos)):
        curso = str(cursos[i]).strip()
        practica = str(practicas[i]).strip().upper().replace(" ", "")

        if i == 0:
            columnas.append("Alumno")
        elif practica.startswith("P"):
            columnas.append(f"{curso}_{practica}")
        else:
            columnas.append(None)

    df.columns = columnas
    df = df.iloc[fila_practicas + 1:].reset_index(drop=True)
    df = df.loc[:, df.columns.notna()]
    df = df[df["Alumno"].notna()]

    return df


# 📄 GENERAR PDF
def generar_pdf(alumno, df, file_id):
    data = df[df["Alumno"] == alumno]

    if data.empty:
        raise Exception("Alumno sin datos")

    file_path = f"{UPLOAD_FOLDER}/{file_id}_{alumno}.pdf"

    styles = getSampleStyleSheet()
    elements = []

    elements.append(NextPageTemplate("TwoCol"))

    elements.append(Paragraph("📊 <b>REPORTE DE PRÁCTICAS CALIFICADAS</b>", styles["Title"]))
    elements.append(Spacer(1, 8))
    elements.append(Paragraph(f"<b>Alumno:</b> {alumno}", styles["Normal"]))
    elements.append(Spacer(1, 15))

    cursos_dict = {}
    promedios = {}

    for col in df.columns:
        if col != "Alumno":
            curso, _ = col.split("_")
            cursos_dict.setdefault(curso, []).append(col)

    for curso, cols in cursos_dict.items():

        cols = sorted(cols, key=lambda x: int(x.split("_")[1].replace("P", "")))

        tabla = [["Práctica", "Nota"]]
        notas = []

        for col in cols:
            val = data.iloc[0][col]

            if pd.isna(val) or val == "":
                tabla.append([col.split("_")[1], "Pendiente"])
            else:
                val = float(val)
                notas.append(val)
                tabla.append([col.split("_")[1], val])

        if notas:
            promedio = round(sum(notas) / len(notas), 2)
        else:
            promedio = 0

        promedios[curso] = promedio

        color_prom = "green" if promedio >= 13 else "red"

        t = Table(tabla)
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.darkgreen),
            ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('GRID', (0,0), (-1,-1), 1, colors.black)
        ]))

        bloque = [
            Paragraph(f"<b>{curso}</b>", styles["Heading3"]),
            t,
            Paragraph(f"<b>Promedio:</b> <font color='{color_prom}'>{promedio}</font>", styles["Normal"]),
            Spacer(1, 12)
        ]

        elements.append(KeepTogether(bloque))

    # 🔥 Ranking en 1 columna
    elements.append(NextPageTemplate("OneCol"))
    elements.append(PageBreak())

    elements.append(Paragraph("🏆 <b>Ranking de Fortalezas</b>", styles["Title"]))
    elements.append(Spacer(1, 15))

    ranking = sorted(promedios.items(), key=lambda x: x[1], reverse=True)

    tabla_rank = [["Posición", "Curso", "Promedio"]]

    for i, (curso, prom) in enumerate(ranking, start=1):
        tabla_rank.append([i, curso, prom])

    t_rank = Table(tabla_rank)
    t_rank.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.darkblue),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('GRID', (0,0), (-1,-1), 1, colors.black)
    ]))

    elements.append(t_rank)

    doc = SimpleDocTemplate(file_path, pagesize=letter)

    frame1 = Frame(doc.leftMargin, doc.bottomMargin, 260, doc.height, id='col1')
    frame2 = Frame(doc.leftMargin + 270, doc.bottomMargin, 260, doc.height, id='col2')
    template_two = PageTemplate(id='TwoCol', frames=[frame1, frame2])

    frame_full = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id='full')
    template_one = PageTemplate(id='OneCol', frames=[frame_full])

    doc.addPageTemplates([template_two, template_one])

    doc.build(elements)

    return file_path


# 🌐 RUTAS
@app.route("/", methods=["GET", "POST"])
def index():
    try:
        if request.method == "POST":
            file = request.files["file"]

            if file.filename == "":
                return render_template_string(HTML, error="Selecciona un archivo")

            file_id = str(uuid.uuid4())
            path = os.path.join(UPLOAD_FOLDER, f"{file_id}.xlsx")
            file.save(path)

            return redirect(url_for("ver_alumnos", file_id=file_id))

        return render_template_string(HTML)

    except Exception as e:
        return render_template_string(HTML, error=str(e))


@app.route("/alumnos/<file_id>")
def ver_alumnos(file_id):
    try:
        path = os.path.join(UPLOAD_FOLDER, f"{file_id}.xlsx")
        df = procesar_excel(path)

        alumnos = df["Alumno"].dropna().unique().tolist()

        return render_template_string(HTML, alumnos=alumnos, file_id=file_id)

    except Exception as e:
        return render_template_string(HTML, error=str(e))


@app.route("/reporte", methods=["POST"])
def reporte():
    try:
        alumno = request.form["alumno"]
        file_id = request.form["file_id"]

        path = os.path.join(UPLOAD_FOLDER, f"{file_id}.xlsx")
        df = procesar_excel(path)

        pdf = generar_pdf(alumno, df, file_id)

        return send_file(pdf, as_attachment=True)

    except Exception as e:
        return f"Error: {str(e)}<br><pre>{traceback.format_exc()}</pre>"


if __name__ == "__main__":
    app.run(debug=True, use_reloader=False)
