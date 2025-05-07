from flask import Flask, render_template, request, send_file
import pandas as pd
from docx import Document
import os
import tempfile
import subprocess

app = Flask(__name__)
df_global = None
UPLOAD_FOLDER = "uploads"
TEMPLATE_PATH = "plantilla/base.docx"

@app.route("/", methods=["GET", "POST"])
def index():
    global df_global
    if request.method == "POST":
        file = request.files["file"]
        filepath = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(filepath)
        df_global = pd.read_excel(filepath)
        records = df_global.to_dict(orient="records")
        return render_template("index.html", data=records)
    return render_template("index.html", data=None)

@app.route("/generar/<int:row_id>")
def generar(row_id):
    global df_global
    if df_global is None or row_id >= len(df_global):
        return "Registro no encontrado", 404

    row = df_global.iloc[row_id].fillna("")

    doc = Document(TEMPLATE_PATH)

    reemplazos = {
        "modalidad": row["modalidad"],
        "tipo_contratacion": row["tipo_contratacion"],
        "numero": str(row["numero"]),
        "identificacion": str(row["identificacion"]),
        "razón_social": row["razón_social"],
        "valor": f"${row['valor']:,}".replace(",", "."),
        "fecha_suscripcion": row["fecha_suscripcion"],
        "fecha_legalizacion": row["fecha_legalizacion"],
        "fecha_cdp": row["fecha_cdp"],
        "numero_cdp": str(row["numero_cdp"]),
        "cuenta_cdp": row["cuenta_cdp"],
        "valor_cdp": f"${row['valor_cdp']:,}".replace(",", "."),
        "fecha_rp": row["fecha_rp"],
        "numero_rp": str(row["numero_rp"]),
        "cuenta_rp": row["cuenta_rp"],
        "valor_rp": f"${row['valor_rp']:,}".replace(",", "."),
        "NOMBRE_SUPERVISOR": row["NOMBRE_SUPERVISOR"],
        "CARGO_SUPERVISOR": row["CARGO_SUPERVISOR"],
    }

    for p in doc.paragraphs:
        for key, val in reemplazos.items():
            if f"«{key}»" in p.text:
                p.text = p.text.replace(f"«{key}»", str(val))

    # Guardar .docx temporal
    tmp_dir = tempfile.gettempdir()
    docx_path = os.path.join(tmp_dir, f"certificado_{row_id}.docx")
    pdf_path = docx_path.replace(".docx", ".pdf")
    doc.save(docx_path)

    # Convertir a PDF usando LibreOffice
    subprocess.run([
        "libreoffice",
        "--headless",
        "--convert-to", "pdf",
        "--outdir", tmp_dir,
        docx_path
    ], check=True)

    return send_file(pdf_path, as_attachment=True, download_name=f"certificado_{row_id+1}.pdf")

if __name__ == "__main__":
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    app.run(host="0.0.0.0", port=8080)
