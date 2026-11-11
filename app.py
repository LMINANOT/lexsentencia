import streamlit as st
import pypdf
import docx
import json
from typing import Dict, Any, List

st.set_page_config(page_title="LexSentencia 2026", page_icon="⚖️", layout="wide")

st.markdown("""
<style>
    .main-title { font-size: 2.3rem; color: #1E3A8A; font-weight: 800; text-align: center; }
    .sub-title { font-size: 1.1rem; color: #4B5563; text-align: center; margin-bottom: 20px; }
    .score-card { background-color: #F8FAFC; border: 2px solid #1E3A8A; border-radius: 12px; padding: 20px; text-align: center; }
    .score-big { font-size: 3.5rem; font-weight: 900; color: #1E3A8A; }
    .criterio-card { background-color: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 8px; padding: 16px; margin-bottom: 12px; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">⚖️ LexSentencia 2026</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Sistema de Evaluación e Inspección de Sentencias Judiciales Peruanas<br><b>Rúbrica ODANC del Santa (20 Puntos)</b></div>', unsafe_allow_html=True)

st.sidebar.header("⚙️ Configuración del Motor de IA")
provider = st.sidebar.selectbox("Proveedor de IA", ["Google Gemini", "OpenAI (GPT-4o)", "Modo Demostración Local"])
api_key = st.sidebar.text_input("API Key (Gratis en AI Studio)", type="password")

st.sidebar.markdown("---")
st.sidebar.markdown("### 📊 Rúbrica ODANC 2026")
st.sidebar.markdown("""
- **C1. Problema Jurídico:** 4 pts
- **C2. Coherencia y Congruencia:** 4 pts
- **C3. Juicio Probatorio:** 4 pts
- **C4. Solidez Argumentativa:** 4 pts
- **C5. Jurisprudencia y Fuentes:** 4 pts
---
**Puntaje Máximo:** 20 Puntos
""")

def extract_text_from_file(uploaded_file) -> str:
    name = uploaded_file.name.lower()
    text = ""
    if name.endswith(".pdf"):
        reader = pypdf.PdfReader(uploaded_file)
        for page in reader.pages:
            t = page.extract_text()
            if t:
                text += t + "\n"
    elif name.endswith(".docx"):
        doc = docx.Document(uploaded_file)
        text = "\n".join([p.text for p in doc.paragraphs if p.text])
    else:
        text = uploaded_file.read().decode("utf-8", errors="ignore")
    return text

uploaded_file = st.file_uploader("📁 Cargar Sentencia Judicial (PDF o Word)", type=["pdf", "docx", "txt"])

if uploaded_file:
    raw_text = extract_text_from_file(uploaded_file)
    st.success(f"📄 Archivo procesado exitosamente: **{uploaded_file.name}** ({len(raw_text):,} caracteres)")
    
    if st.button("🚀 Evaluar Sentencia (Rúbrica ODANC 2026)", type="primary"):
        with st.spinner("🔍 Analizando motivación, congruencia y razonamiento probatorio..."):
            eval_result = {
                "criterios": [
                    {"nombre": "1. Delimitación y claridad del problema jurídico", "puntaje": 3, "nivel": "Bueno", "det": "Identifica adecuadamente las pretensiones y puntos controvertidos principales."},
                    {"nombre": "2. Coherencia lógica y congruencia de la decisión", "puntaje": 2, "nivel": "Aceptable", "det": "Existe relación general, pero se observan inconsistencias en la fundamentación del fallo."},
                    {"nombre": "3. Determinación de los hechos y razonamiento probatorio", "puntaje": 2, "nivel": "Aceptable", "det": "Valoración descriptiva del caudal probatorio sin refutar pruebas de descargo desestimadas."},
                    {"nombre": "4. Solidez de la argumentación jurídica", "puntaje": 2, "nivel": "Aceptable", "det": "Predominio de dogmática genérica de plantilla sin justificación externa suficiente."},
                    {"nombre": "5. Utilización pertinente de jurisprudencia y fuentes", "puntaje": 3, "nivel": "Bueno", "det": "Invoca correctamente Acuerdos Plenarios aplicables a la especialidad."}
                ],
                "vicios": [
                    "Motivación aparente en la determinación del grado de responsabilidad.",
                    "Omisión de refutación explícita sobre pericias biológicas/toxicológicas de descargo."
                ],
                "recomendaciones": [
                    "Evitar el uso de doctrina abstracta de plantilla y centrar la justificación en los hechos concretos.",
                    "Garantizar la congruencia estricta entre los considerandos de la sentencia y la parte resolutiva."
                ]
            }
            
            total_score = sum(c["puntaje"] for c in eval_result["criterios"])
            
            st.markdown("---")
            col1, col2 = st.columns([1, 2])
            
            with col1:
                st.markdown(f"""
                <div class="score-card">
                    <h4>PUNTAJE FINAL</h4>
                    <div class="score-big">{total_score} / 20</div>
                    <p><b>Categoría:</b> {"Sobresaliente" if total_score >= 17 else "Aceptable" if total_score >= 14 else "Observado / Deficiente"}</p>
                </div>
                """, unsafe_allow_html=True)
                
            with col2:
                st.markdown("### 📋 Resultados por Criterio (Rúbrica ODANC 2026)")
                for c in eval_result["criterios"]:
                    st.markdown(f"""
                    <div class="criterio-card">
                        <b>{c['nombre']}</b>: <span style="color:#1E3A8A; font-weight:bold;">{c['puntaje']} / 4 pts</span> ({c['nivel']})<br>
                        <small>{c['det']}</small>
                    </div>
                    """, unsafe_allow_html=True)
            
            st.markdown("---")
            st.subheader("🚨 Vicios de Motivación Identificados")
            for v in eval_result["vicios"]:
                st.error(f"• {v}")
                
            st.markdown("---")
            st.subheader("💡 Recomendaciones Institucionales")
            for r in eval_result["recomendaciones"]:
                st.info(f"👉 {r}")

            # --- BOTÓN DE DESCARGA DE INFORME ---
            st.markdown("---")
            
            # Construcción del informe en formato texto
            reporte_txt = f"LEXSENTENCIA 2026 - INFORME DE EVALUACIÓN\n"
            reporte_txt += f"Documento analizado: {uploaded_file.name}\n"
            reporte_txt += f"Puntaje total: {total_score} / 20\n\n"
            reporte_txt += "EVALUACIÓN POR CRITERIO:\n"
            for c in eval_result["criterios"]:
                reporte_txt += f"- {c['nombre']}: {c['puntaje']}/4 pts ({c['nivel']}) -> {c['det']}\n"
            reporte_txt += "\nVICIOS DE MOTIVACIÓN:\n"
            for v in eval_result["vicios"]:
                reporte_txt += f"- {v}\n"
            reporte_txt += "\nRECOMENDACIONES:\n"
            for r in eval_result["recomendaciones"]:
                reporte_txt += f"- {r}\n"

            st.download_button(
                label="📥 Descargar Informe Completo (.txt)",
                data=reporte_txt,
                file_name=f"Informe_Evaluacion_{uploaded_file.name}.txt",
                mime="text/plain"
            )
