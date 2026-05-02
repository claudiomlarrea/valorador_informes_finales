# =============================================================================
# Plantilla institucional UCCuyo (Streamlit) — checklist al portar otra repo:
#   • _resolve_escudo_path (+ assets/escudo_uccuyo.png commitado) y/o URL en
#     secrets/environment (UCC_ESCUDO_URL → HTTPS; st.image sí lo muestra).
#   • Encabezado: st.container(horizontal=True, key=...) + st.image / markdown
#   • CSS .st-key-… , .ucci-inst-banner-text , .ucci-main-title-card
#   • .streamlit/config.toml [theme] (secondaryBackgroundColor del cargador)
# =============================================================================
import os

import streamlit as st
import pandas as pd
import pdfplumber
import yaml
import io
from pathlib import Path
from docx import Document
from docx.shared import Pt
from datetime import datetime
from openpyxl import Workbook

_APP_DIR = Path(__file__).resolve().parent


def _resolve_escudo_path() -> Path | None:
    """Ruta del escudo si existe en assets/ (debe estar en el repo para Streamlit Cloud)."""
    assets = _APP_DIR / "assets"
    if not assets.is_dir():
        return None
    for name in ("escudo_uccuyo.png", "escudo_uccuyo.jpg", "escudo_uccuyo.jpeg"):
        p = assets / name
        if p.is_file():
            return p
    return None


def _extras_escudo_url() -> str | None:
    """URL HTTPS opcional (Streamlit suele sanitizar data: en Markdown)."""
    v = os.environ.get("UCC_ESCUDO_URL", "").strip()
    if v:
        return v

    try:
        secrets = st.secrets
    except Exception:
        return None

    get = getattr(secrets, "get", None)
    if callable(get):
        for key in ("UCC_ESCUDO_URL", "ucc_escudo_url"):
            raw = get(key)
            if isinstance(raw, str) and raw.strip():
                return raw.strip()

    for key in ("ucc_escudo_url", "UCC_ESCUDO_URL"):
        try:
            raw = secrets[key]
        except Exception:
            continue
        if isinstance(raw, str) and raw.strip():
            return raw.strip()

    return None


def _resolved_escudo_for_st_image():
    """Ruta local (repo) o URL para st.image."""
    path = _resolve_escudo_path()
    if path is not None:
        return path

    url = _extras_escudo_url()
    if url:
        return url

    return None


st.set_page_config(layout="wide")

st.markdown(
    """
<style>
:root {
    --ucci-green: #00664d;
    --ucci-green-dark: #00523e;
    --ucci-accent: #28a745;
    --ucci-page-bg: #f0f2f6;
    --ucci-sidebar-bg: #262730;
    --ucci-text: #262730;
}

.stApp {
    background-color: var(--ucci-page-bg);
}

.block-container {
    padding-top: 1.25rem;
}

section[data-testid="stSidebar"] {
    background-color: var(--ucci-sidebar-bg);
}
[data-testid="stSidebar"] [data-testid="stMarkdown"],
[data-testid="stSidebar"] span,
[data-testid="stSidebar"] label {
    color: rgba(255, 255, 255, 0.92);
}

/* —— Banner institucional: logo con st.image + textos Markdown (clave contenedor horizontal) —— */
/* Streamlit slug del key=\"ucci_inst_banner\" → clase st-key-ucci-inst-banner */
div[data-testid="stVerticalBlockBorderWrapper"]:has([class*="st-key-ucci-inst-banner"]),
div[data-testid="stVerticalBlockBorderWrapper"]:has([class*="st-key-ucci_inst_banner"]) {
    background: var(--ucci-green) !important;
    border-radius: 12px !important;
    margin-bottom: 1.25rem !important;
    padding: 0.85rem 1.35rem !important;
    border: none !important;
    box-sizing: border-box !important;
}

div[data-testid="stVerticalBlockBorderWrapper"]:has([class*="st-key-ucci-inst-banner"]) img,
div[data-testid="stVerticalBlockBorderWrapper"]:has([class*="st-key-ucci_inst_banner"]) img {
    display: block;
    max-height: 96px;
    width: auto !important;
    max-width: 104px !important;
    height: auto !important;
    object-fit: contain;
    background: rgba(255, 255, 255, 0.98);
    border-radius: 10px;
    padding: 0.35rem;
    box-sizing: border-box;
}

[class*="st-key-ucci-inst-banner"],
[class*="st-key-ucci_inst_banner"] {
    min-width: 0;
}

.ucci-inst-banner-text {
    flex: 1 1 auto;
    min-width: 0;
}

.ucci-inst-banner-text,
.ucci-inst-banner-text * {
    color: #ffffff !important;
}

.ucci-inst-banner-text a {
    color: #ffffff !important;
    text-decoration: underline dotted rgba(255, 255, 255, 0.55);
}

.ucci-inst-banner-text sup,
.ucci-inst-banner-text code {
    color: inherit !important;
}

.header-ucciuyo h1.ucci-banner-heading,
.header-ucciuyo h2.ucci-banner-heading,
.header-ucciuyo h3.ucci-banner-heading {
    color: #ffffff !important;
    margin: 0;
    line-height: 1.2;
    font-family: "Source Sans Pro", ui-sans-serif, system-ui, sans-serif;
}

.header-ucciuyo h1.ucci-banner-heading {
    font-size: clamp(1.35rem, 2.8vw, 1.95rem);
    font-weight: 700;
}

.header-ucciuyo h2.ucci-banner-heading {
    margin-top: 0.5rem !important;
    font-size: clamp(1rem, 2vw, 1.22rem);
    font-weight: 500;
}

.header-ucciuyo h3.ucci-banner-heading {
    margin-top: 0.3rem !important;
    font-size: clamp(0.85rem, 1.4vw, 1rem);
    font-weight: 400;
    color: rgba(255, 255, 255, 0.93) !important;
}

/* Título de la app tipo tarjeta */
.ucci-main-title-card {
    background: #ffffff;
    border-radius: 12px;
    padding: 1.1rem 1.35rem 1rem;
    margin-bottom: 1rem;
    box-shadow: 0 1px 3px rgba(0, 0, 0, 0.07);
    border: 1px solid rgba(0, 38, 28, 0.08);
    box-sizing: border-box;
}

.ucci-main-title-card h1.ucci-app-title {
    margin: 0 !important;
    padding: 0 !important;
    font-size: clamp(1.35rem, 2.6vw, 1.72rem);
    font-weight: 700;
    color: var(--ucci-text) !important;
}

.ucci-main-title-card p.ucci-app-subtitle {
    margin: 0.55rem 0 0 0 !important;
    font-size: 0.95rem;
    line-height: 1.45;
    color: #5f6368 !important;
}

/* Contenido general */
h1:not(.ucci-banner-heading):not(.ucci-app-title),
h2:not(.ucci-banner-heading),
h3:not(.ucci-banner-heading),
h4 {
    color: var(--ucci-green-dark) !important;
}
p,
label {
    color: var(--ucci-text) !important;
}

/* Carga de archivo: tarjeta clara + franja oscura viene del tema (secondaryBackgroundColor) */
[data-testid="stFileUploader"] {
    background-color: #ffffff !important;
    border-radius: 12px !important;
    padding: 0.95rem !important;
    border: 1px dashed rgba(0, 102, 77, 0.3) !important;
    margin-top: 0.15rem !important;
}

[data-testid="stFileUploader"] button[kind],
[data-testid="stFileUploader"] button {
    background-color: var(--ucci-green) !important;
    color: white !important;
    border-radius: 8px;
    border: none !important;
}

.stButton > button {
    background-color: var(--ucci-green) !important;
    color: white !important;
    border-radius: 8px;
    border: none;
    font-weight: 600;
}
.stButton > button:hover {
    background-color: var(--ucci-green-dark) !important;
    border-color: transparent !important;
}

[data-testid="stDownloadButton"] button {
    background-color: var(--ucci-green) !important;
    color: white !important;
    border-radius: 8px;
    border: none;
    font-weight: 600;
}
[data-testid="stDownloadButton"] button:hover {
    background-color: var(--ucci-green-dark) !important;
    border-color: transparent !important;
}

div[data-testid="stAlert"] {
    border-radius: 10px;
}

[data-baseweb="slider"] {
    color: var(--ucci-green);
}

.stButton button span,
[data-testid="stDownloadButton"] button span {
    color: white !important;
}

.stButton > button,
.stButton > button * {
    color: white !important;
}

[data-testid="stDownloadButton"] button,
[data-testid="stDownloadButton"] button * {
    color: white !important;
}

[data-testid="stFileUploader"] button span,
[data-testid="stFileUploader"] button div,
[data-testid="stFileUploader"] button p {
    color: white !important;
}

.stSlider label,
[data-testid="stTextInput"] label,
[data-testid="stFileUploader"] label {
    position: relative;
    padding-left: 1rem;
}
.stSlider label::before,
[data-testid="stTextInput"] label::before,
[data-testid="stFileUploader"] label::before {
    content: "";
    position: absolute;
    left: 0;
    top: 0.45rem;
    width: 9px;
    height: 9px;
    border-radius: 50%;
    background: var(--ucci-accent);
}
</style>
""",
    unsafe_allow_html=True,
)

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
_esc_img = _resolved_escudo_for_st_image()

# Logo: st.image (archivo local o HTTPS). Markdown con <img data:…> suele filtrarlo el sanitizer de Streamlit.
with st.container(
    horizontal=True,
    key="ucci_inst_banner",
    horizontal_alignment="left",
    vertical_alignment="center",
    gap="small",
):
    if _esc_img is not None:
        st.image(_esc_img, width=104, use_container_width=False)
    st.markdown(
        """
<div class="ucci-inst-banner-text header-ucciuyo">
<h1 class="ucci-banner-heading">Universidad Católica de Cuyo</h1>
<h2 class="ucci-banner-heading">Secretaría de Investigación</h2>
<h3 class="ucci-banner-heading">Consejo de Investigación</h3>
</div>
""",
        unsafe_allow_html=True,
    )

st.markdown(
    """
<div class="ucci-main-title-card">
<h1 class="ucci-app-title">📊 Valorador de Informes Finales</h1>
<p class="ucci-app-subtitle">Subí un informe final (PDF o DOCX) para evaluarlo automáticamente según la rúbrica institucional.</p>
</div>
""",
    unsafe_allow_html=True,
)

archivo = st.file_uploader("Cargar archivo", type=["pdf", "docx"])

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
