from flask import Flask, request, render_template_string, send_file
import pandas as pd
import os
import uuid
import traceback
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet

app = Flask(__name__)
app.config["PROPAGATE_EXCEPTIONS"] = True

UPLOAD_FOLDER = "temp"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

HTML = """
<!DOCTYPE html>
<html>
<head>
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
    border-radius: 8px;
    border: none;
    background: #4CAF50;
    color: white;
}
.error {
    color: red;
    font-weight: bold;
}
</style>
</head>
<body>

<div class="box">
<h2>📊 Subir Excel</h2>
<form method="POST" enctype="multipart/form-data">
<input type="file" name="file" required>
<button>Procesar</button>
</form>
</div>

{% if error %}
<div class="box error">
{{error}}
</div>
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

# 🔥 PROCESAR EXCEL (ULTRA ROBUSTO)
def procesar_excel(path):
    df = pd.read_excel(path, header=None)

    # 🔍 detectar columna de alumno
    col_alumno = None
    for col in df.columns:
        valores = df[col].astype(str)
        if valores.str.contains("ALUMNO|NOMBRE", case=False).any():
            col_alumno = col
            break

    if col_alumno is None:
        col_alumno = 0  # fallback

    # encabezados
    cursos = df.iloc[0].ffill()
    practicas = df.iloc[1]

    columnas = []

    for i in range(len(cursos)):
        curso = str(cursos[i]).strip()
        practica = str(practicas[i]).strip()

        if i == col_alumno:
            columnas.append("Alumno")
        elif practica.upper().startswith("P"):
            columnas.append(f"{curso}_{practica}")
        else:
            columnas.append(None)

    df.columns = columnas

    # datos
    df = df.iloc[2:].reset_index(drop=True)

    # limpiar
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

    elements.append(Paragraph("Reporte de Prácticas", styles["Title"]))
    elements.append(Spacer(1, 10))
    elements.append(Paragraph(f"Alumno: {alumno}", styles["Normal"]))
    elements.append(Spacer(1, 20))

    cursos_dict = {}

    # agrupar cursos
    for col in df.columns:
        if col != "Alumno" and "_" in col:
            curso, practica = col.split("_")
            cursos_dict.setdefault(curso, []).append(col)

    # generar tablas
    for curso, cols in cursos_dict.items():

        # ordenar P1 → P6
        try:
            cols = sorted(cols, key=lambda x: int(x.split("_")[1].replace("P","")))
        except:
            pass

        tabla = [["Práctica", "Nota"]]
        notas = []

        for col in cols:
            val = data.iloc[0][col]

            try:
                val = float(val)
            except:
                val = 0

            tabla.append([col.split("_")[1], val])
            notas.append(val)

        promedio = round(sum(notas)/len(notas), 2) if notas else 0

        elements.append(Paragraph(f"<b>{curso}</b>", styles["Heading2"]))

        t = Table(tabla)
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.green),
            ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('GRID', (0,0), (-1,-1), 1, colors.black)
        ]))

        elements.append(t)
        elements.append(Spacer(1, 10))
        elements.append(Paragraph(f"Promedio: {promedio}", styles["Normal"]))
        elements.append(Spacer(1, 15))

    doc.build(elements)
    return file_path


# 🌐 RUTAS
@app.route("/", methods=["GET", "POST"])
def index():
    try:
        if request.method == "POST":
            file = request.files["file"]

            file_id = str(uuid.uuid4())
            path = f"{UPLOAD_FOLDER}/{file_id}.xlsx"
            file.save(path)

            df = procesar_excel(path)

            alumnos = df["Alumno"].dropna().unique().tolist()

            return render_template_string(HTML, alumnos=alumnos, file_id=file_id)

        return render_template_string(HTML)

    except Exception as e:
        return render_template_string(HTML, error=str(e))


@app.route("/reporte", methods=["POST"])
def reporte():
    try:
        alumno = request.form["alumno"]
        file_id = request.form["file_id"]

        path = f"{UPLOAD_FOLDER}/{file_id}.xlsx"
        df = procesar_excel(path)

        pdf = generar_pdf(alumno, df, file_id)

        return send_file(pdf, as_attachment=True)

    except Exception as e:
        return f"Error: {str(e)}<br><pre>{traceback.format_exc()}</pre>"


if __name__ == "__main__":
    app.run(debug=True)
