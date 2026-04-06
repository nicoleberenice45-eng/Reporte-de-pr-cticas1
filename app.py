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

# 🎨 fondo degradado
def dibujar_fondo(canvas, doc):
    width, height = letter

    # color base verde claro
    canvas.setFillColorRGB(0.85, 1, 0.85)
    canvas.rect(0, 0, width, height, fill=1, stroke=0)

    # franja superior amarilla
    canvas.setFillColorRGB(1, 0.9, 0.3)
    canvas.rect(0, height - 100, width, 100, fill=1, stroke=0)


# ✅ PROCESAR EXCEL (detecta filas correctas)
def procesar_excel(path):
    df = pd.read_excel(path, header=None)

    df = df.dropna(how="all").reset_index(drop=True)

    # detectar fila de prácticas
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


# 📄 GENERAR PDF (2 COLUMNAS)
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

    for col in df.columns:
        if col != "Alumno":
            curso, _ = col.split("_")
            cursos_dict.setdefault(curso, []).append(col)

    cursos_lista = list(cursos_dict.items())
    mitad = len(cursos_lista) // 2

    col1 = cursos_lista[:mitad]
    col2 = cursos_lista[mitad:]

    def crear_bloque(cursos):
        bloque = []
        for curso, cols in cursos:

            cols = sorted(cols, key=lambda x: int(x.split("_")[1].replace("P", "")))

            tabla = [["Práctica", "Nota"]]

            for col in cols:
                val = data.iloc[0][col]
                tabla.append([col.split("_")[1], val])

            t = Table(tabla)
            t.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.green),
                ('TEXTCOLOR', (0,0), (-1,0), colors.white),
                ('GRID', (0,0), (-1,-1), 1, colors.black)
            ]))

            bloque.append(Paragraph(f"<b>{curso}</b>", styles["Heading3"]))
            bloque.append(t)
            bloque.append(Spacer(1, 10))

        return bloque

    bloque1 = crear_bloque(col1)
    bloque2 = crear_bloque(col2)

    tabla_final = Table([[bloque1, bloque2]], colWidths=[270, 270])

    elements.append(tabla_final)

    doc.build(elements, onFirstPage=dibujar_fondo, onLaterPages=dibujar_fondo)

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
