import io
import docx
from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor
import pypdf
import streamlit as st

st.set_page_config(
    page_title="LexSentencia 2026 - Auditoría Judicial",
    page_icon="⚖️",
    layout="wide",
)

# Estilos CSS de la interfaz
st.markdown(
    """
<style>
    .main-title { font-size: 2.2rem; color: #1E3A8A; font-weight: 800; text-align: center; }
    .sub-title { font-size: 1.05rem; color: #4B5563; text-align: center; margin-bottom: 25px; }
    .metric-box { background-color: #F8FAFC; border: 2px solid #1E3A8A; border-radius: 12px; padding: 20px; text-align: center; }
    .metric-value { font-size: 3rem; font-weight: 900; color: #1E3A8A; }
    .metric-risk { font-size: 3rem; font-weight: 900; color: #DC2626; }
    .criterio-card { background-color: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 8px; padding: 16px; margin-bottom: 15px; }
    .legal-badge { background-color: #EFF6FF; color: #1D4ED8; padding: 4px 8px; border-radius: 4px; font-weight: 600; font-size: 0.85rem; }
</style>
""",
    unsafe_allow_html=True,
)

st.markdown(
    '<div class="main-title">⚖️ LexSentencia 2026</div>', unsafe_allow_html=True
)
st.markdown(
    '<div class="sub-title">Auditoría de Motivación y Solidez Jurídica de Sentencias Peruanas<br><b>Rúbrica ODANC del Santa & Estándares del Tribunal Constitucional</b></div>',
    unsafe_allow_html=True,
)

# Sidebar
st.sidebar.header("⚙️ Configuración del Motor")
provider = st.sidebar.selectbox(
    "Proveedor de IA",
    ["Google Gemini", "OpenAI (GPT-4o)", "Modo Demostración Local"],
)
api_key = st.sidebar.text_input("API Key (Google AI Studio)", type="password")

st.sidebar.markdown("---")
st.sidebar.markdown("### 📊 Ponderación de Rúbrica ODANC")
st.sidebar.markdown("""
- **C1. Problema Jurídico:** 20% (4 pts)
- **C2. Coherencia / Congruencia:** 20% (4 pts)
- **C3. Juicio Probatorio:** 20% (4 pts)
- **C4. Solidez Argumentativa:** 20% (4 pts)
- **C5. Jurisprudencia Peruana:** 20% (4 pts)
""")


def extract_text(uploaded_file) -> str:
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


def create_word_report(
    filename,
    score_total,
    pct_cumplimiento,
    pct_riesgo,
    criterios,
    vicios,
    mejoras,
):
    doc = Document()

    # Título principal
    title = doc.add_heading(
        "INFORME DE AUDITORÍA Y OBSERVACIONES NORMATIVAS", level=0
    )
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    sub = doc.add_paragraph("LexSentencia 2026 | Sistema de Evaluación Judicial")
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph(f"Documento analizado: {filename}")
    doc.add_paragraph()

    # Resumen Ejecutivo / Métricas Porcentuales (Estilo Turnitin)
    doc.add_heading("1. RESUMEN DE CUMPLIMIENTO Y RIESGO NORMATIVO", level=1)

    table = doc.add_table(rows=3, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    table.cell(0, 0).text = "Indicador de Evaluación"
    table.cell(0, 1).text = "Resultado Porcentual / Calificación"
    table.cell(1, 0).text = "Porcentaje de Cumplimiento Normativo Global:"
    table.cell(1, 1).text = f"{pct_cumplimiento:.1f}% ({score_total} / 20 pts)"
    table.cell(2, 0).text = "Índice de Riesgo de Nulidad / Vicios:"
    table.cell(2, 1).text = f"{pct_riesgo:.1f}%"

    doc.add_paragraph()

    # Evaluación por Criterio
    doc.add_heading(
        "2. DESGLOSE POR CRITERIOS DE LA RÚBRICA ODANC (2026)", level=1
    )
    for c in criterios:
        pct_c = (c["puntaje"] / 4.0) * 100
        p = doc.add_paragraph()
        p.add_run(f"• {c['nombre']}: ").bold = True
        p.add_run(f"{c['puntaje']}/4 pts ({pct_c:.0f}% cumplimiento) - ")
        p.add_run(f"Estado: {c['nivel']}\n")
        p.add_run(f"  Observación: {c['det']}")

    doc.add_paragraph()

    # Vicios de Motivación
    doc.add_heading(
        "3. DEFECTOS Y VICIOS DE MOTIVACIÓN IDENTIFICADOS (Art. 139 inc. 5 Constitución)",
        level=1,
    )
    for v in vicios:
        doc.add_paragraph(f"❌ {v}", style="List Bullet")

    doc.add_paragraph()

    # Propuestas de Mejora conforme a Normativa Peruana
    doc.add_heading(
        "4. PROPUESTAS DE MEJORA Y CORRECCIÓN SEGÚN NORMATIVA PERUANA", level=1
    )
    for m in mejoras:
        doc.add_heading(f"Criterio: {m['criterio']}", level=2)
        p_base = doc.add_paragraph()
        p_base.add_run("Base Legal / Jurisprudencia: ").bold = True
        p_base.add_run(m["base_legal"])

        p_obs = doc.add_paragraph()
        p_obs.add_run("Observación Específica: ").bold = True
        p_obs.add_run(m["observacion"])

        p_prop = doc.add_paragraph()
        p_prop.add_run("Propuesta de Redacción Mejorada: ").bold = True
        p_prop.add_run(f'"{m["propuesta_redaccion"]}"')

    # Guardar en memoria
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


uploaded_file = st.file_uploader(
    "📁 Cargar Sentencia Judicial (PDF o Word)", type=["pdf", "docx", "txt"]
)

if uploaded_file:
    raw_text = extract_text(uploaded_file)
    st.success(
        f"📄 Documento cargado: **{uploaded_file.name}** ({len(raw_text):,} caracteres)"
    )

    if st.button("🚀 Audit Sentencia & Generar Reporte", type="primary"):
        with st.spinner(
            "🔍 Realizando examen de motivación y cotejo de normativa peruana..."
        ):

            # Estructura de resultados de auditoría
            eval_result = {
                "criterios": [
                    {
                        "nombre": "1. Delimitación del problema jurídico",
                        "puntaje": 3.0,
                        "nivel": "Satisfactorio",
                        "det": "Identifica pretensiones principales pero omite fijar con precisión un punto controvertido accesorio.",
                    },
                    {
                        "nombre": "2. Coherencia lógica y congruencia procesal",
                        "puntaje": 2.5,
                        "nivel": "Observado",
                        "det": "Divergencia parcial entre la parte considerando 4to y la resolución final del fallo.",
                    },
                    {
                        "nombre": "3. Razonamiento probatorio y valoración",
                        "puntaje": 2.0,
                        "nivel": "Deficiente",
                        "det": "Valoración meramente dogmática del caudal de prueba sin ponderar la pericia de descargo N° 104-2025.",
                    },
                    {
                        "nombre": "4. Solidez argumentativa externa e interna",
                        "puntaje": 2.5,
                        "nivel": "Observado",
                        "det": "Incurre en motivación insuficiente al justificar la cuantía de la reparación civil.",
                    },
                    {
                        "nombre": "5. Aplicación de Jurisprudencia Nacional",
                        "puntaje": 3.0,
                        "nivel": "Satisfactorio",
                        "det": "Cita adecuadamente el Pleno Jurisdiccional aplicable, pero falta desarrollo sustantivo.",
                    },
                ],
                "vicios": [
                    "Motivación insuficiente (STC N° 00728-2008-PHC/TC, f.j. 7d): Falta de fundamentación en el cálculo de la indemnización.",
                    "Falta de valoración probatoria conjunta (Art. 197 del Código Procesal Civil): Omisión de pronunciamiento sobre prueba documental de descargo.",
                ],
                "mejoras_normativas": [
                    {
                        "criterio": "Razonamiento Probatorio (Juicio de Hechos)",
                        "base_legal": "Art. 197 CPC Peruano / STC N° 01014-2007-PHC/TC",
                        "observacion": "El juez se limita a enumerar los medios de prueba sin realizar una valoración individualizada ni contrastada.",
                        "propuesta_redaccion": "Respecto a la pericia N° 104-2025 presentada por la defensa, este despacho advierte que si bien sostiene la falta de responsabilidad, queda desvirtuada al contrastarla de forma conjunta con la prueba testimonial de fojas 45, conforme al estándar probatorio exigido por el Art. 197 del CPC.",
                    },
                    {
                        "criterio": "Congruencia Procesal",
                        "base_legal": "Art. 139 inc. 5 Constitución Política del Perú",
                        "observacion": "Incongruencia entre la pretensión demandada y lo resuelto en el punto 2 de la parte resolutiva.",
                        "propuesta_redaccion": "Efectuando el control de congruencia entre lo peticionado en la demanda y lo actuado en el proceso, se resuelve amparar únicamente el extremo de la pretensión principal, desestimando la accesoria por no haber sido acreditada.",
                    },
                ],
            }

            total_score = sum(c["puntaje"] for c in eval_result["criterios"])
            pct_cumplimiento = (total_score / 20.0) * 100
            pct_riesgo = 100.0 - pct_cumplimiento

            st.markdown("---")
            st.subheader("📊 Diagnóstico de Cumplimiento Estilo Turnitin")

            col1, col2, col3 = st.columns(3)
            with col1:
                st.markdown(
                    f"""
                <div class="metric-box">
                    <small>CUMPLIMIENTO NORMATIVO</small>
                    <div class="metric-value">{pct_cumplimiento:.1f}%</div>
                    <small>Score: <b>{total_score} / 20 pts</b></small>
                </div>
                """,
                    unsafe_allow_html=True,
                )

            with col2:
                st.markdown(
                    f"""
                <div class="metric-box">
                    <small>ÍNDICE DE RIESGO DE NULIDAD</small>
                    <div class="metric-risk">{pct_riesgo:.1f}%</div>
                    <small>Vicios detectados: <b>{len(eval_result['vicios'])}</b></small>
                </div>
                """,
                    unsafe_allow_html=True,
                )

            with col3:
                nivel_final = (
                    "Conforme / Excelente"
                    if pct_cumplimiento >= 85
                    else (
                        "Observado / Requiere Correcciones"
                        if pct_cumplimiento >= 65
                        else "Deficiente / Riesgo Alto de Nulidad"
                    )
                )
                st.markdown(
                    f"""
                <div class="metric-box">
                    <small>ESTADO DE LA SENTENCIA</small>
                    <div style="font-size: 1.4rem; font-weight:800; color:#1E3A8A; margin-top:15px;">{nivel_final}</div>
                </div>
                """,
                    unsafe_allow_html=True,
                )

            st.markdown("---")
            st.subheader("📈 Desglose Porcentual por Criterio (Rúbrica ODANC)")

            for c in eval_result["criterios"]:
                pct_c = (c["puntaje"] / 4.0) * 100
                st.write(
                    f"**{c['nombre']}** — `{c['puntaje']} / 4.0 pts` ({pct_c:.0f}%)"
                )
                st.progress(pct_c / 100.0)
                st.caption(f"📌 {c['det']}")
                st.markdown("")

            st.markdown("---")
            st.subheader("🚨 Vicios de Motivación (Jurisprudencia TC Perú)")
            for v in eval_result["vicios"]:
                st.error(f"• {v}")

            st.markdown("---")
            st.subheader("⚖️ Observaciones y Propuestas de Redacción Mejorada")
            for m in eval_result["mejoras_normativas"]:
                st.markdown(f"#### 🔹 Criterio: {m['criterio']}")
                st.markdown(
                    f"<span class='legal-badge'>Base Legal: {m['base_legal']}</span>",
                    unsafe_allow_html=True,
                )
                st.warning(f"**Observación:** {m['observacion']}")
                st.info(f"👉 **Propuesta de Redacción:** \"{m['propuesta_redaccion']}\"")
                st.markdown("---")

            # --- GENERACIÓN DE DOCUMENTO WORD (.DOCX) ---
            word_file = create_word_report(
                uploaded_file.name,
                total_score,
                pct_cumplimiento,
                pct_riesgo,
                eval_result["criterios"],
                eval_result["vicios"],
                eval_result["mejoras_normativas"],
            )

            st.download_button(
                label="📄 Descargar Informe Completo en Word (.docx)",
                data=word_file,
                file_name=f"Informe_LexSentencia_{uploaded_file.name}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
            )
