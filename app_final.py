import streamlit as st
import pandas as pd
import numpy as np
import pdfplumber
import yaml
import io
from docx import Document
from docx.shared import Pt
from datetime import datetime
from openpyxl import Workbook

# ==============================
# CARGA DE CONFIGURACIÓN
# ==============================
with open("rubric_final.yaml", "r", encoding="utf-8") as f:
    config = yaml.safe_load(f)

weights = config["weights"]
thresholds = config["thresholds"]
keywords = config["keywords"]

# ==============================
# FUNCIONES DE PROCESAMIENTO
# ==============================
def extract_text(file):
    """Extrae texto desde PDF o DOCX"""
    if file.name.endswith(".pdf"):
        text = ""
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                text += page.extract_text() + "\n"
        return text
    elif file.name.endswith(".docx"):
        doc = Document(file)
        return "\n".join([p.text for p in doc.paragraphs])
    else:
        return ""

def auto_score(text, keywords_dict):
    """Calcula puntaje automático según palabras clave"""
    scores = {}
    for section, keys in keywords_dict.items():
        found = sum(k.lower() in text.lower() for k in keys)
        # Escala simple: 0 a 4
        scores[section] = min(4, found)
    return scores

def weighted_score(scores, weights):
    """Calcula el puntaje total ponderado"""
    total = sum(scores[s] * weights[s] for s in scores)
    max_total = sum(weights.values()) * 4
    percent = (total / max_total) * 100
    return percent

def generate_excel(scores, percent, thresholds):
    """Genera archivo Excel con resultados"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Resultados"
    ws.append(["Criterio", "Puntaje (0-4)"])
    for k, v in scores.items():
        ws.append([k, v])
    ws.append([])
    ws.append(["Puntaje total (%)", round(percent, 2)])
    if percent >= thresholds["aprobado"]:
        result = "Aprobado"
    elif percent >= thresholds["aprobado_obs"]:
        result = "Aprobado con observaciones"
    else:
        result = "No aprobado"
    ws.append(["Dictamen", result])

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def generate_word(scores, percent, thresholds):
    """Genera dictamen en Word (sin apartado de evidencia analizada)"""
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Arial"
    style.font.size = Pt(11)

    doc.add_heading("Valoración del Informe Final", level=1)

    # Fecha y encabezado
    doc.add_paragraph(f"Fecha de evaluación: {datetime.today().strftime('%d/%m/%Y')}")
    doc.add_paragraph("")

    # PUNTAJES
    doc.add_heading("Puntajes por criterio", level=2)
    table = doc.add_table(rows=1, cols=2)
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "Criterio"
    hdr_cells[1].text = "Puntaje (0–4)"

    for k, v in scores.items():
        row_cells = table.add_row().cells
        row_cells[0].text = k.replace("_", " ").capitalize()
        row_cells[1].text = str(v)

    doc.add_paragraph("")
    total_text = f"Puntaje total: {round(percent,2)}%"
    doc.add_paragraph(total_text)

    # DICTAMEN FINAL
    if percent >= thresholds["aprobado"]:
        result = "Aprobado"
    elif percent >= thresholds["aprobado_obs"]:
        result = "Aprobado con observaciones"
    else:
        result = "No aprobado"

    doc.add_heading("Dictamen final", level=2)
    doc.add_paragraph(result)

    # Se eliminó por completo el bloque de evidencia analizada

    # Observaciones
    doc.add_heading("Observaciones del evaluador", level=2)
    doc.add_paragraph("................................................................................")
    doc.add_paragraph("................................................................................")
    doc.add_paragraph("................................................................................")

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# ==============================
# INTERFAZ STREAMLIT
# ==============================
st.title("📊 Valorador de Informes Finales")
st.write("Subí un informe final en formato PDF o DOCX para evaluarlo automáticamente según la rúbrica institucional.")

uploaded_file = st.file_uploader("Cargar archivo", type=["pdf", "docx"])

if uploaded_file:
    text = extract_text(uploaded_file)

    with st.expander("Ver texto extraído"):
        st.text_area("Texto completo", text, height=300)

    st.subheader("Evaluación automática")
    auto_scores = auto_score(text, keywords)
    df = pd.DataFrame(auto_scores.items(), columns=["Criterio", "Puntaje (0–4)"])
    st.dataframe(df, use_container_width=True)

    percent = weighted_score(auto_scores, weights)
    st.metric(label="Puntaje total (%)", value=round(percent, 2))

    if percent >= thresholds["aprobado"]:
        result = "✅ Aprobado"
    elif percent >= thresholds["aprobado_obs"]:
        result = "⚠️ Aprobado con observaciones"
    else:
        result = "❌ No aprobado"
    st.success(f"Dictamen automático: {result}")

    st.subheader("Ajuste manual (opcional)")
    manual_scores = {}
    for k in auto_scores.keys():
        manual_scores[k] = st.slider(f"{k.replace('_',' ').capitalize()}", 0, 4, int(auto_scores[k]))

    if st.button("Generar informes"):
        final_percent = weighted_score(manual_scores, weights)
        excel_file = generate_excel(manual_scores, final_percent, thresholds)
        word_file = generate_word(manual_scores, final_percent, thresholds)

        st.download_button("⬇️ Descargar Excel", excel_file, file_name="valoracion_informe_final.xlsx")
        st.download_button("⬇️ Descargar Word", word_file, file_name="valoracion_informe_final.docx")

        st.success("Archivos generados correctamente (sin apartado de evidencia analizada).")
