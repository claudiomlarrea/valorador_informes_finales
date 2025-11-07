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
import os

# ============================
# CONFIGURACIÓN
# ============================
with open("rubric_final.yaml", "r", encoding="utf-8") as f:
    config = yaml.safe_load(f)

weights    = config["weights"]
thresholds = config["thresholds"]
keywords   = config["keywords"]
scale_min  = config["scale"]["min"]
scale_max  = config["scale"]["max"]

# ============================
# FUNCIONES
# ============================
def extract_text(file):
    """Extrae texto desde PDF o DOCX"""
    if file.name.lower().endswith(".pdf"):
        text = ""
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                text += (page.extract_text() or "") + "\n"
        return text
    elif file.name.lower().endswith(".docx"):
        doc = Document(file)
        return "\n".join(p.text for p in doc.paragraphs)
    else:
        return ""

def auto_score(text, keywords_dict):
    """Calcula puntajes automáticos según palabras clave (simple)"""
    scores = {}
    text_low = (text or "").lower()
    for section, keys in keywords_dict.items():
        found = sum((k or "").lower() in text_low for k in keys)
        scores[section] = min(scale_max, int(found))
    return scores

def weighted_score(scores, weights):
    """Calcula el puntaje total ponderado (0–100)"""
    total = sum(scores[s] * weights[s] for s in scores)
    max_total = sum(weights.values()) * scale_max
    return (total / max_total) * 100 if max_total > 0 else 0.0

def generate_excel(scores, percent, thresholds):
    """Genera archivo Excel con resultados"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Resultados"
    ws.append(["Criterio", "Puntaje (0–4)"])
    for k, v in scores.items():
        ws.append([k.replace("_", " ").capitalize(), v])
    ws.append([])
    ws.append(["Puntaje total (%)", round(percent, 2)])
    if percent >= thresholds["aprobado"]:
        result = "Aprobado"
    elif percent >= thresholds["aprobado_obs"]:
        result = "Aprobado con observaciones"
    else:
        result = "No aprobado"
    ws.append(["Dictamen", result])
    out = io.BytesIO()
    wb.save(out); out.seek(0)
    return out

def generate_word(scores, percent, thresholds, nombre_proyecto=""):
    """Genera dictamen Word incluyendo el nombre del proyecto"""
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Arial"
    style.font.size = Pt(11)

    base_title = "UCCuyo – Valoración del Informe Final"
    np_clean = (nombre_proyecto or "").strip()
    if np_clean:
        doc.add_heading(f'{base_title} "Del proyecto {np_clean}"', level=1)
    else:
        doc.add_heading(base_title, level=1)

    doc.add_paragraph(f"Fecha de evaluación: {datetime.today().strftime('%Y-%m-%d %H:%M')}")
    doc.add_paragraph("")

    # Resultados
    doc.add_heading("Puntajes por criterio", level=2)
    table = doc.add_table(rows=1, cols=2)
    hdr = table.rows[0].cells
    hdr[0].text = "Criterio"
    hdr[1].text = "Puntaje (0–4)"
    for k, v in scores.items():
        row = table.add_row().cells
        row[0].text = k.replace("_", " ").capitalize()
        row[1].text = str(v)

    doc.add_paragraph(f"\nPuntaje total: {round(percent, 2)}%")

    # Dictamen final
    if percent >= thresholds["aprobado"]:
        result = "Aprobado"
    elif percent >= thresholds["aprobado_obs"]:
        result = "Aprobado con observaciones"
    else:
        result = "No aprobado"

    doc.add_heading("Dictamen final", level=2)
    doc.add_paragraph(result)

    # Observaciones
    doc.add_heading("Observaciones del evaluador", level=2)
    doc.add_paragraph("................................................................................")
    doc.add_paragraph("................................................................................")
    doc.add_paragraph("................................................................................")

    out = io.BytesIO()
    doc.save(out); out.seek(0)
    return out

def filename_without_ext(name: str) -> str:
    """Nombre del archivo sin extensión, para usar como fallback del proyecto."""
    base = os.path.basename(name or "")
    return os.path.splitext(base)[0]

# ============================
# INTERFAZ STREAMLIT
# ============================
st.title("📘 Valorador de Informes Finales")
st.write("Subí un informe final (PDF o DOCX) para evaluarlo automáticamente según la rúbrica institucional.")

uploaded_file = st.file_uploader("Cargar informe (PDF o DOCX)", type=["pdf", "docx"])

# Campo SIEMPRE visible: nombre del proyecto
default_name = filename_without_ext(uploaded_file.name) if uploaded_file else ""
nombre_proyecto = st.text_input("Nombre del proyecto (aparecerá en el Word):", value=default_name)

if uploaded_file:
    text = extract_text(uploaded_file)
    with st.expander("Ver texto extraído"):
        st.text_area("Texto completo", text, height=300)

    st.subheader("Evaluación automática")
    auto_scores = auto_score(text, keywords)
    df = pd.DataFrame([(k.replace("_"," ").capitalize(), v) for k,v in auto_scores.items()],
                      columns=["Criterio", "Puntaje (0–4)"])
    st.dataframe(df, use_container_width=True)

    percent_auto = weighted_score(auto_scores, weights)
    st.metric(label="Puntaje total (%)", value=round(percent_auto, 2))

    dictamen_auto = ("Aprobado" if percent_auto >= thresholds["aprobado"]
                     else "Aprobado con observaciones" if percent_auto >= thresholds["aprobado_obs"]
                     else "No aprobado")
    st.success(f"Dictamen automático: {dictamen_auto}")

    st.subheader("Ajuste manual (opcional)")
    manual_scores = {}
    for k in auto_scores.keys():
        manual_scores[k] = st.slider(f"{k.replace('_',' ').capitalize()}", scale_min, scale_max, int(auto_scores[k]))

    # Por defecto exporta con automáticos (evita diferencias por sliders guardados)
    use_auto = st.checkbox("Generar informe con los puntajes automáticos (recomendado)", value=True)

    if st.button("Generar informes"):
        scores_to_use = auto_scores if use_auto else manual_scores
        percent_final = weighted_score(scores_to_use, weights)
        excel_file = generate_excel(scores_to_use, percent_final, thresholds)
        word_file  = generate_word(scores_to_use, percent_final, thresholds, nombre_proyecto)

        st.download_button("⬇️ Descargar Excel", excel_file, file_name="valoracion_informe_final.xlsx")
        st.download_button("⬇️ Descargar Word",  word_file,  file_name="valoracion_informe_final.docx")

        st.success("Informe generado con puntajes {}."
                   .format("automáticos" if use_auto else "ajustados manualmente"))
