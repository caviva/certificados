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

# Función segura para formatear montos
def formato_moneda(valor):
    try:
        return f"${float(valor):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return str(valor)

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
        "modalidad": row.get("modalidad", ""),
        "tipo_contratacion": row.get("tipo_contratacion", ""),
        "numero": str(row.get("numero", "")),
        "identificacion": str(row.get("identificacion", "")),
        "razón_social": row.get("razón_social", ""),
        "valor": formato_moneda(row.get("valor", 0)),
        "fecha_suscripcion": row.get("fecha_suscripcion", ""),
        "fecha_legalizacion": row.get("fecha_legalizacion", ""),
        "fecha_cdp": row.get("fecha_cdp", ""),
        "numero_cdp": str(row.get("numero_cdp", "")),
        "cuenta_cdp": row.get("cuenta_cdp", ""),
        "valor_cdp": formato_moneda(row.get("valor_cdp", 0)),
        "fecha_rp": row.get("fecha_rp", ""),
        "numero_rp": str(row.get("numero_rp", "")),
        "cuenta_rp": row.get("cuenta_rp", ""),
        "valor_rp": formato_moneda(row.get("valor_rp", 0)),
        "NOMBRE_SUPERVISOR": row.get("NOMBRE_SUPERVISOR", ""),
        "CARGO_SUPERVISOR": row.get("CARGO_SUPERVISOR", "")
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
