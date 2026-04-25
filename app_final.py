import streamlit as st
import pandas as pd
import pdfplumber
import yaml
import io
from docx import Document
from docx.shared import Pt
from datetime import datetime
from openpyxl import Workbook

st.set_page_config(layout="wide")
st.markdown("""
<style>

/* FONDO GENERAL */
.stApp {
    background-color: #E6E6E6;
}

/* HEADER */

/* TÍTULOS */
.header-uccuyo h1,
.header-uccuyo h2,
.header-uccuyo h3 {
    color: white !important;
}

/* TEXTO */
p, label, span {
    color: #1a1a1a;
}

/* CAJA UPLOAD */
[data-testid="stFileUploader"] {
    background-color: white;
    border-radius: 10px;
    padding: 15px;
}

/* BOTONES */
.stButton button,
[data-testid="stDownloadButton"] button,
[data-testid="stFileUploader"] button {
    background-color: #064a3f !important;
    color: white !important;
    border-radius: 8px;
    border: none;
    font-weight: 600;
}

.stButton button:hover {
    background-color: #0B6B5D !important;
}

</style>
""", unsafe_allow_html=True)

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


# 🔥 FUNCIÓN INSTITUCIONAL FINAL
def auto_score(text, keywords_dict):
    scores = {}
    text_low = (text or "").lower()

    # ============================
    # DETECCIONES GLOBALES
    # ============================

    tiene_objetivos = "objetivo" in text_low
    tiene_resultados = "resultado" in text_low
    tiene_datos = "tabla" in text_low or "fig" in text_low or "datos" in text_low
    tiene_porcentajes = "%" in text_low
    tiene_transferencia = "publicación" in text_low or "congreso" in text_low
    tiene_rrhh = "tesis" in text_low or "beca" in text_low

    for section, keys in keywords_dict.items():

        found = sum(k in text_low for k in keys)

        # BASE EXIGENTE
        if found >= 6:
            base = 4
        elif found >= 4:
            base = 3
        elif found >= 2:
            base = 2
        elif found >= 1:
            base = 1
        else:
            base = 0

        bonus = 0
        penalty = 0

        # OBJETIVOS
        if section == "objetivos":
            if tiene_porcentajes:
                bonus += 1
            if "cumpl" in text_low or "logr" in text_low:
                bonus += 1
            if not tiene_objetivos:
                penalty += 3

        # CRONOGRAMA
        if section == "cronograma":
            if tiene_porcentajes:
                bonus += 1
            else:
                penalty += 3

        # RESULTADOS
        if section == "resultados":
            if tiene_datos:
                bonus += 1
            else:
                penalty += 3

        # RRHH
        if section == "formacion_rrhh":
            if tiene_rrhh:
                bonus += 1
            else:
                penalty += 3

        # TRANSFERENCIA
        if section == "transferencia":
            if tiene_transferencia:
                bonus += 1
            else:
                penalty += 3

        # CALIDAD FORMAL
        if section == "calidad_formal":
            if "bibliografía" not in text_low and "citación" not in text_low:
                penalty += 2

        # IMPACTO
        if section == "impacto":
            if "impacto" not in text_low:
                penalty += 2

        # COHERENCIA GLOBAL
        if tiene_objetivos and not tiene_resultados:
            penalty += 3

        if tiene_resultados and not tiene_datos:
            penalty += 2

        score = base + bonus - penalty
        scores[section] = max(0, min(4, score))

    return scores


def weighted_score(scores, weights):
    total = sum(scores[s] * weights[s] for s in scores)
    max_total = sum(weights.values()) * 4

    percent = (total / max_total) * 100

    # PENALIZACIÓN GLOBAL
    criterios_bajos = sum(1 for s in scores.values() if s <= 1)

    if criterios_bajos >= 4:
        percent -= 10

    if criterios_bajos >= 6:
        percent -= 15

    # PROMEDIO GENERAL
    promedio = sum(scores.values()) / len(scores)

    if promedio < 1.5:
        percent -= 15

    return max(0, percent)


def generate_excel(scores, percent, thresholds):
    wb = Workbook()
    ws = wb.active
    ws.title = "Resultados"

    ws.append(["Criterio", "Puntaje"])
    for k, v in scores.items():
        ws.append([k, v])

    ws.append([])
    ws.append(["Total (%)", round(percent, 2)])

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

    doc.add_heading("Resultados", 2)

    table = doc.add_table(rows=1, cols=2)
    hdr = table.rows[0].cells
    hdr[0].text = "Criterio"
    hdr[1].text = "Puntaje"

    for k, v in scores.items():
        row = table.add_row().cells
        row[0].text = k
        row[1].text = str(v)

    doc.add_paragraph(f"\nTotal: {round(percent, 2)}%")

    if percent >= thresholds["aprobado"]:
        result = "Aprobado"
    elif percent >= thresholds["aprobado_obs"]:
        result = "Aprobado con observaciones"
    else:
        result = "No aprobado"

    doc.add_heading("Dictamen", 2)
    doc.add_paragraph(result)

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# ============================
# INTERFAZ
# ============================

st.markdown(
"""<div class="header-uccuyo" style="background: linear-gradient(90deg, #0b5d4b, #177e6c); padding: 30px; border-radius: 15px; margin: 0 auto 30px auto; max-width: 900px;">

<h1 style="color: white !important; margin:0; font-size:40px; font-weight:700;">
Universidad Católica de Cuyo
</h1>

<h2 style="color: white !important; margin-top:10px; font-size:22px; font-weight:500;">
Secretaría de Investigación
</h2>

<h3 style="color: #d6f2ec !important; margin-top:5px; font-size:18px; font-weight:400;">
Consejo de Investigación
</h3>

</div>""",
unsafe_allow_html=True
)

st.title("📊 Valorador de Informes Finales")

archivo = st.file_uploader("Subir informe", type=["pdf", "docx"])

if archivo:
    texto = extract_text(archivo)

    scores = auto_score(texto, keywords)

    df = pd.DataFrame(scores.items(), columns=["Criterio", "Puntaje"])
    st.dataframe(df)

    percent = weighted_score(scores, weights)

    st.metric("Resultado (%)", round(percent, 2))

    if percent >= thresholds["aprobado"]:
        resultado = "Aprobado"
    elif percent >= thresholds["aprobado_obs"]:
        resultado = "Aprobado con observaciones"
    else:
        resultado = "No aprobado"

    st.success(f"Dictamen: {resultado}")

    nombre = st.text_input("Nombre del proyecto")

    if st.button("Generar informe"):
        excel = generate_excel(scores, percent, thresholds)
        word = generate_word(scores, percent, thresholds, nombre)

        st.download_button("Descargar Excel", excel, "resultado.xlsx")
        st.download_button("Descargar Word", word, "resultado.docx")
