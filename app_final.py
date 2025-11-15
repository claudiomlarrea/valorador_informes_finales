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

# ============================
# CONFIGURACIÓN
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
    else:
        return ""

def auto_score(text, keywords_dict):
    """Calcula puntajes automáticos según palabras clave"""
    scores = {}
    text_low = (text or "").lower()
    for section, keys in keywords_dict.items():
        found = sum((k or "").lower() in text_low for k in keys)
        scores[section] = min(4, found)
    return scores

def weighted_score(scores, weights):
    """Calcula el puntaje total ponderado (%) a partir de puntajes 0–4"""
    total = sum(scores[s] * weights[s] for s in scores)
    max_total = sum(weights.values()) * 4
    percent = (total / max_total) * 100 if max_total > 0 else 0.0
    return percent

def generate_excel(scores, percent, thresholds):
    """Genera archivo Excel con resultados"""
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
    """Genera dictamen Word incluyendo el nombre del proyecto"""
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Arial"
    style.font.size = Pt(11)

    # Encabezado
    base_title = "UCCuyo – Valoración de Informe Final"
    nombre_clean = (nombre_proyecto or "").strip()
    if nombre_clean:
        doc.add_heading(f'{base_title} "Del proyecto {nombre_clean}"', level=1)
    else:
        doc.add_heading(base_title, level=1)

    doc.add_paragraph(f"Fecha: {datetime.today().strftime('%Y-%m-%d %H:%M')}")
    doc.add_paragraph("")

    # Puntajes por criterio
    doc.add_heading("Resultados por criterio", level=2)
    table = doc.add_table(rows=1, cols=2)
    hdr = table.rows[0].cells
    hdr[0].text = "Criterio"
    hdr[1].text = "Puntaje (0–4)"
    for k, v in scores.items():
        row = table.add_row().cells
        row[0].text = k.replace("_", " ").capitalize()
        row[1].text = str(v)

    percent_text = f"\nCumplimiento: {round(percent, 2)}%"
    doc.add_paragraph(percent_text)

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
    doc.add_paragraph("..............................................................................")
    doc.add_paragraph("..............................................................................")
    doc.add_paragraph("..............................................................................")

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# ============================
# INTERFAZ STREAMLIT
# ============================
st.title("📗 Valorador de Informes Finales")
st.write("Subí un informe final (PDF o DOCX) para evaluarlo automáticamente según la rúbrica institucional.")

uploaded_file = st.file_uploader("Cargar archivo", type=["pdf", "docx"])

if uploaded_file:
    text = extract_text(uploaded_file)

    with st.expander("Ver texto extraído"):
        st.text_area("Texto completo", text, height=300)

    # --- Evaluación automática (referencia) ---
    st.subheader("Evaluación automática")
    auto_scores = auto_score(text, keywords)
    df = pd.DataFrame(auto_scores.items(), columns=["Criterio", "Puntaje (0–4)"])
    st.dataframe(df, use_container_width=True)

    auto_percent = weighted_score(auto_scores, weights)
    st.metric(label="Puntaje automático inicial (%)", value=round(auto_percent, 2))

    # --- Ajuste manual ---
    st.subheader("Ajuste manual (opcional)")
    manual_scores = {}
    for k in auto_scores.keys():
        manual_scores[k] = st.slider(
            f"{k.replace('_',' ').capitalize()}",
            0,
            4,
            int(auto_scores[k]),
        )

    # Puntaje total AJUSTADO (este es el que importa)
    adjusted_percent = weighted_score(manual_scores, weights)
    st.metric(label="Puntaje total ajustado (%)", value=round(adjusted_percent, 2))

    # Dictamen con ajuste manual
    if adjusted_percent >= thresholds["aprobado"]:
        result = "✅ Aprobado"
    elif adjusted_percent >= thresholds["aprobado_obs"]:
        result = "⚠️ Aprobado con observaciones"
    else:
        result = "❌ No aprobado"
    st.success(f"Dictamen (con ajuste manual): {result}")

    # Nombre del proyecto para el Word
    nombre_proyecto = st.text_input("Nombre del proyecto (aparecerá en el Word):", "")

    # Generar informes SIEMPRE con los valores ajustados
    if st.button("Generar informes"):
        final_percent = adjusted_percent
        excel_file = generate_excel(manual_scores, final_percent, thresholds)
        word_file = generate_word(manual_scores, final_percent, thresholds, nombre_proyecto)

        st.download_button(
            "⬇️ Descargar Excel",
            excel_file,
            file_name="valoracion_informe_final.xlsx",
        )
        st.download_button(
            "⬇️ Descargar Word",
            word_file,
            file_name="valoracion_informe_final.docx",
        )

        st.success("Informe generado con los puntajes ajustados manualmente.")
