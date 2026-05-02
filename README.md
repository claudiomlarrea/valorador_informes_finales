# Valorador de **Informes Finales** (Streamlit)

Calculadora institucional para **valorar informes finales** de proyectos de investigación a partir de **PDF o DOCX**.
- Carga el archivo
- Extrae el texto automáticamente
- Evalúa 11 criterios (0–4) con **ponderaciones configurables** (`rubric_final.yaml`)
- Permite **ajuste manual** por el evaluador
- Exporta **Excel** y **Word** con dictamen

## Umbrales (definitivos)
- **≥ 60% → Aprobado**
- **50–59.99% → Aprobado con observaciones**
- **< 50% → No aprobado**

## Ejecutar
```bash
pip install -r requirements.txt
streamlit run app_final.py
```

## Archivos
- `app_final.py` — Aplicación Streamlit (no compara con proyecto original)
- `rubric_final.yaml` — Pesos/umbral y palabras clave
- `requirements.txt` — Dependencias
- `runtime.txt` — Versión de Python para Streamlit Cloud
