import streamlit as st
import pandas as pd
import pdfplumber
import yaml
import io
from docx import Document
from docx.shared import Pt
from datetime import datetime
from openpyxl import Workbook

# ============================
# CARGA DE RÚBRICA
# ============================
with open("rubric_final.yaml", "r", encoding="utf-8") as f:
    config = yaml.safe_load(f)

weights = config["weights"]
thresholds = config["thresholds"]
keywords = config["keywords"]

# ============================
# FUNCIONES
# ============================

def extract_text(file):
    """Extrae texto desde PDF o DOCX"""
    if file.name.endswith(".pdf"):
        text = ""
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                text += (page.extract_text() or "") + "\n"
        return text

    elif file.name.endswith(".docx"):
        doc = Document(file)
        return "\n".join([p.text for p in doc.paragraphs])

    return ""


# 🔥 NUEVA FUNCIÓN INTELIGENTE
def auto_score(text, keywords_dict):
    scores = {}
    text_low = (text or "").lower()

    for section, keys in keywords_dict.items():
        found = sum(k in text_low for k in keys)

        bonus = 0

        # OBJETIVOS → detectar cumplimiento real
        if section == "objetivos":
            if "%" in text_low:
                bonus += 1
            if "cumpl" in text_low or "logr" in text_low:
                bonus += 1

        # CRONOGRAMA → detectar avance
        if section == "cronograma":
            if "avance" in text_low or "%" in text_low:
                bonus += 1

        # RESULTADOS → detectar evidencia
        if section == "resultados":
            if "tabla" in text_low or "fig" in text_low:
                bonus += 1

        # RRHH → detectar formación real
        if section == "formacion_rrhh":
            if "tesis" in text_low or "beca" in text_low:
                bonus += 1

        scores[section] = min(4, found + bonus)

    return scores


def weighted_score(scores, weights):
    total = sum(scores[s] * weights[s] for s in scores)
    max_total = sum(weights.values()) * 4
    percent = (total / max_total) * 100 if max_total > 0 else 0.0
    return percent


def generate_excel(scores, percent, thresholds):
    wb = Workbook()
    ws = wb.active
    ws.title = "Resultados"

    ws.append(["Criterio", "Puntaje (0–4)"])
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


def generate_word(scores, percent, thresholds, nombre_proyecto=""):
    doc = Document()

    style = doc.styles["Normal"]
    style.font.name = "Arial"
    style.font.size = Pt(11)

    doc.add_heading("Valoración de Informe Final", 1)
    doc.add_paragraph(f"Fecha: {datetime.today().strftime('%Y-%m-%d %H:%M')}")

    if nombre_proyecto:
        doc.add_paragraph(f"Proyecto: {nombre_proyecto}")

    doc.add_heading("Resultados por criterio", 2)

    table = doc.add_table(rows=1, cols=2)
    hdr = table.rows[0].cells
    hdr[0].text = "Criterio"
    hdr[1].text = "Puntaje"

    for k, v in scores.items():
        row = table.add_row().cells
        row[0].text = k.replace("_", " ").capitalize()
        row[1].text = str(v)

    doc.add_paragraph(f"\nCumplimiento total: {round(percent, 2)}%")

    if percent >= thresholds["aprobado"]:
        result = "Aprobado"
    elif percent >= thresholds["aprobado_obs"]:
        result = "Aprobado con observaciones"
    else:
        result = "No aprobado"

    doc.add_heading("Dictamen final", 2)
    doc.add_paragraph(result)

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output


# ============================
# INTERFAZ
# ============================

st.title("📗 Valorador de Informes Finales")

uploaded_file = st.file_uploader("Subir informe (PDF o DOCX)", type=["pdf", "docx"])

if uploaded_file:
    text = extract_text(uploaded_file)

    st.subheader("Texto extraído")
    st.text_area("Contenido", text, height=300)

    st.subheader("Evaluación automática")

    auto_scores = auto_score(text, keywords)

    df = pd.DataFrame(auto_scores.items(), columns=["Criterio", "Puntaje"])
    st.dataframe(df)

    percent = weighted_score(auto_scores, weights)

    st.metric("Puntaje total (%)", round(percent, 2))

    if percent >= thresholds["aprobado"]:
        result = "Aprobado"
    elif percent >= thresholds["aprobado_obs"]:
        result = "Aprobado con observaciones"
    else:
        result = "No aprobado"

    st.success(f"Dictamen: {result}")

    nombre = st.text_input("Nombre del proyecto")

    if st.button("Generar informes"):
        excel = generate_excel(auto_scores, percent, thresholds)
        word = generate_word(auto_scores, percent, thresholds, nombre)

        st.download_button("Descargar Excel", excel, "resultado.xlsx")
        st.download_button("Descargar Word", word, "resultado.docx")
