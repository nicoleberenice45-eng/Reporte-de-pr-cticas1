from flask import Flask, request, render_template_string, send_file
import pandas as pd
import os
import uuid
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet

app = Flask(__name__)

UPLOAD_FOLDER = "temp"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# 🎨 HTML
HTML = """
<!DOCTYPE html>
<html>
<head>
    <title>Reportes</title>
    <style>
        body {
            font-family: Arial;
            background: linear-gradient(135deg, #4CAF50, #FFEB3B);
            text-align: center;
        }
        .box {
            background: white;
            padding: 20px;
            margin: 40px auto;
            width: 500px;
            border-radius: 15px;
        }
        button {
            padding: 10px;
            margin: 5px;
            width: 80%;
            border: none;
            border-radius: 8px;
            background: #4CAF50;
            color: white;
            cursor: pointer;
        }
    </style>
</head>
<body>

<div class="box">
    <h2>📊 Subir Excel</h2>
    <form method="POST" enctype="multipart/form-data">
        <input type="file" name="file" required>
        <button type="submit">Procesar</button>
    </form>
</div>

{% if alumnos %}
<div class="box">
    <h3>Selecciona un alumno</h3>
    {% for alumno in alumnos %}
        <form action="/reporte" method="POST">
            <input type="hidden" name="alumno" value="{{alumno}}">
            <input type="hidden" name="file_id" value="{{file_id}}">
            <button type="submit">{{alumno}}</button>
        </form>
    {% endfor %}
</div>
{% endif %}

</body>
</html>
"""

# 🔍 PROCESAR TU EXCEL REAL
def procesar_excel(path):
    df = pd.read_excel(path, header=None)

    headers_main = df.iloc[5]
    headers_sub = df.iloc[6]

    columnas = []
    for i in range(len(headers_main)):
        if pd.notna(headers_sub[i]):
            columnas.append(str(headers_sub[i]))
        else:
            columnas.append(str(headers_main[i]))

    df.columns = columnas
    df = df.iloc[7:].reset_index(drop=True)

    df.rename(columns={"APELLIDOS Y NOMBRES DEL ALUMNO": "Alumno"}, inplace=True)
    df = df[df["Alumno"].notna()]

    return df

# 📄 PDF
def generar_pdf(alumno, df, file_id):
    data = df[df["Alumno"] == alumno]

    file_path = f"{UPLOAD_FOLDER}/{file_id}_{alumno}.pdf"

    doc = SimpleDocTemplate(file_path, pagesize=letter)
    styles = getSampleStyleSheet()
    elements = []

    elements.append(Paragraph("Reporte de Prácticas", styles["Title"]))
    elements.append(Spacer(1, 10))
    elements.append(Paragraph(f"Alumno: {alumno}", styles["Normal"]))
    elements.append(Spacer(1, 20))

    practicas = ["1º s", "2º s", "3º s", "4º s", "5º s", "6º s"]

    tabla = [["Prácticas"] + practicas + ["Promedio"]]

    for _, row in data.iterrows():
        notas = []
        for p in practicas:
            val = row.get(p, 0)
            try:
                val = float(val)
            except:
                val = 0
            notas.append(val)

        promedio = round(sum(notas)/len(notas), 2)
        tabla.append(["Notas"] + notas + [promedio])

    table = Table(tabla)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.green),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('GRID', (0,0), (-1,-1), 1, colors.black)
    ]))

    elements.append(table)
    doc.build(elements)

    return file_path

# 🌐 RUTAS
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files["file"]

        file_id = str(uuid.uuid4())
        path = f"{UPLOAD_FOLDER}/{file_id}.xlsx"
        file.save(path)

        df = procesar_excel(path)

        alumnos = df["Alumno"].unique().tolist()

        return render_template_string(HTML, alumnos=alumnos, file_id=file_id)

    return render_template_string(HTML, alumnos=None)

@app.route("/reporte", methods=["POST"])
def reporte():
    alumno = request.form["alumno"]
    file_id = request.form["file_id"]

    path = f"{UPLOAD_FOLDER}/{file_id}.xlsx"
    df = procesar_excel(path)

    pdf_path = generar_pdf(alumno, df, file_id)

    return send_file(pdf_path, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
