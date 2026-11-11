import io
import json
import docx
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import RGBColor
import pypdf
import streamlit as st

try:
    import google.generativeai as genai
except ImportError:
    genai = None

st.set_page_config(
    page_title="LexSentencia 2026 - Auditoría Judicial",
    page_icon="⚖️",
    layout="wide",
)

# Estilos CSS de la interfaz (Estilo Turnitin)
st.markdown(
    """
<style>
    .main-title { font-size: 2.2rem; color: #1E3A8A; font-weight: 800; text-align: center; }
    .sub-title { font-size: 1.05rem; color: #4B5563; text-align: center; margin-bottom: 25px; }
    .metric-box { background-color: #F8FAFC; border: 2px solid #1E3A8A; border-radius: 12px; padding: 18px; text-align: center; }
    .metric-value { font-size: 2.8rem; font-weight: 900; color: #1E3A8A; }
    .metric-risk { font-size: 2.8rem; font-weight: 900; color: #DC2626; }
    
    .turnitin-obs-box {
        background-color: #FEF2F2;
        border-left: 5px solid #EF4444;
        padding: 15px;
        border-radius: 6px;
        margin-top: 15px;
        margin-bottom: 5px;
    }
    .turnitin-fix-box {
        background-color: #ECFDF5;
        border-left: 5px solid #10B981;
        padding: 15px;
        border-radius: 6px;
        margin-bottom: 20px;
    }
    .legal-badge { background-color: #DBEAFE; color: #1E40AF; padding: 3px 8px; border-radius: 4px; font-weight: 700; font-size: 0.82rem; }
</style>
""",
    unsafe_allow_html=True,
)

st.markdown('<div class="main-title">⚖️ LexSentencia 2026</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="sub-title">Auditoría de Motivación Judicial & Cotejo Normativo Peruano<br><b>Análisis Párrafo por Párrafo con Google Gemini (Estilo Turnitin)</b></div>',
    unsafe_allow_html=True,
)

# --- CONFIGURACIÓN DE BARRA LATERAL ---
st.sidebar.header("⚙️ Motor de IA & API")
provider = st.sidebar.selectbox("Proveedor de IA", ["Google Gemini (IA Real)", "Modo Demostración Local"])

if provider == "Google Gemini (IA Real)":
    api_key = st.sidebar.text_input("API Key de Google Gemini", type="password", value="")
else:
    api_key = ""
    st.sidebar.info("💡 El Modo Demostración muestra datos de prueba sin consumir API Key.")

st.sidebar.markdown("---")
st.sidebar.header("🛠️ Configuración de Rúbrica Dinámica")

if "rubrica_criterios" not in st.session_state:
    st.session_state.rubrica_criterios = [
        {"Criterio": "1. Delimitación del problema jurídico", "Max_Pts": 4.0},
        {"Criterio": "2. Coherencia lógica y congruencia procesal", "Max_Pts": 4.0},
        {"Criterio": "3. Razonamiento probatorio y valoración", "Max_Pts": 4.0},
        {"Criterio": "4. Solidez argumentativa interna y externa", "Max_Pts": 4.0},
        {"Criterio": "5. Utilización pertinente de jurisprudencia peruana", "Max_Pts": 4.0},
    ]

edited_rubrica = st.sidebar.data_editor(
    st.session_state.rubrica_criterios,
    num_rows="dynamic",
    column_config={
        "Criterio": st.column_config.TextColumn("Nombre del Criterio", width="large"),
        "Max_Pts": st.column_config.NumberColumn("Puntos Máx.", min_value=1.0, max_value=20.0, step=1.0),
    },
    key="rubrica_editor",
)

max_score_posible = sum(item["Max_Pts"] for item in edited_rubrica if item.get("Max_Pts"))
st.sidebar.caption(f"**Puntaje Máximo Total Configurado:** `{max_score_posible:.1f} Puntos`")


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


def create_word_report(filename, score_total, max_score, pct_cumplimiento, pct_riesgo, criterios, hallazgos_parrafos):
    doc = Document()

    title = doc.add_heading("INFORME DE AUDITORÍA JUDICIAL Y REVISIÓN PÁRRAFO POR PÁRRAFO", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    sub = doc.add_paragraph("LexSentencia 2026 | Estándares de Motivación TC Perú & ODANC")
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph(f"Documento Auditado: {filename}")
    doc.add_paragraph()

    doc.add_heading("1. RESUMEN EJECUTIVO Y MÉTRICAS ESTILO TURNITIN", level=1)
    t_summary = doc.add_table(rows=3, cols=2)
    t_summary.style = 'Table Grid'
    t_summary.cell(0, 0).text = "Métrica de Evaluación"
    t_summary.cell(0, 1).text = "Resultado Obtenido"
    t_summary.cell(1, 0).text = "Porcentaje de Cumplimiento Normativo:"
    t_summary.cell(1, 1).text = f"{pct_cumplimiento:.1f}% ({score_total:.1f} / {max_score:.1f} pts)"
    t_summary.cell(2, 0).text = "Índice de Riesgo de Nulidad por Defectos:"
    t_summary.cell(2, 1).text = f"{pct_riesgo:.1f}%"

    doc.add_paragraph()

    doc.add_heading("2. EVALUACIÓN DE RÚBRICA PERSONALIZADA", level=1)
    for c in criterios:
        p = doc.add_paragraph()
        p.add_run(f"• {c['nombre']}: ").bold = True
        p.add_run(f"{c['score']} / {c['max_pts']} pts ({c['pct']:.0f}%) — Estado: {c['nivel']}\n")
        p.add_run(f"  Detalle: {c['detalle']}")

    doc.add_paragraph()

    doc.add_heading("3. AUDITORÍA PÁRRAFO POR PÁRRAFO: OBSERVACIONES Y SUBORDINACIÓN DE TEXTO", level=1)
    for i, h in enumerate(hallazgos_parrafos, 1):
        doc.add_heading(f"Observación N° {i}: {h['criterio_afectado']}", level=2)
        
        p_obs = doc.add_paragraph()
        p_obs.add_run("TEXTO ORIGINAL OBSERVADO EN LA SENTENCIA:\n").bold = True
        run_obs = p_obs.add_run(f'"{h["texto_original"]}"')
        run_obs.font.italic = True
        run_obs.font.color.rgb = RGBColor(180, 0, 0)
        
        p_law = doc.add_paragraph()
        p_law.add_run("Vicio Detectado / Base Legal: ").bold = True
        p_law.add_run(f"{h['vicio']} | Sustento: {h['base_legal']}")

        p_fix = doc.add_paragraph()
        p_fix.add_run("PROPUESTA DE REDACCIÓN SUBSANATORIA SUGERIDA:\n").bold = True
        run_fix = p_fix.add_run(f'"{h["propuesta_redaccion"]}"')
        run_fix.font.bold = True
        run_fix.font.color.rgb = RGBColor(0, 100, 0)

        doc.add_paragraph("─" * 40)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


uploaded_file = st.file_uploader("📁 Cargar Sentencia Judicial (PDF o Word)", type=["pdf", "docx", "txt"])

if uploaded_file:
    raw_text = extract_text(uploaded_file)
    st.success(f"📄 Documento cargado: **{uploaded_file.name}** ({len(raw_text):,} caracteres)")

    if st.button("🚀 Audit Sentencia (Modo Turnitin + Rúbrica Personalizada)", type="primary"):
        with st.spinner("🔍 Conectando con Google Gemini y auditando la sentencia párrafo por párrafo..."):
            
            criterios_res = []
            hallazgos_res = []
            ejecutado_con_ia = False
            model_used_name = ""

            rubrica_str = "\n".join([f"- {item['Criterio']}: max {item['Max_Pts']} pts" for item in edited_rubrica])

            prompt_text = (
                "Actúa como Auditor Jurídico de Sentencias Peruanas (Art. 139 inc. 5 Constitución, STC 00728-2008-PHC/TC).\n"
                "Evalúa el texto adjunto conforme a esta RÚBRICA PERSONALIZADA:\n"
                + rubrica_str + "\n\n"
                "Instrucciones:\n"
                "1. Otorga un puntaje a cada criterio sin exceder su max_pts.\n"
                "2. Extrae extractos reales de la sentencia que contengan vicios de motivación y redacta la versión subsanada.\n\n"
                "Responde ÚNICAMENTE en JSON con este formato:\n"
                "{\n"
                '  "criterios_evaluados": [\n'
                '    {"nombre": "Nombre del Criterio", "score": 3.0, "max_pts": 4.0, "nivel": "Satisfactorio", "detalle": "Justificación"}\n'
                "  ],\n"
                '  "hallazgos_parrafos": [\n'
                "    {\n"
                '      "criterio_afectado": "Nombre del Criterio",\n'
                '      "texto_original": "Párrafo o frase exacta observada en la sentencia",\n'
                '      "vicio": "Vicio de motivación identificado",\n'
                '      "base_legal": "Norma o jurisprudencia peruana aplicable",\n'
                '      "propuesta_redaccion": "Propuesta de redacción corregida"\n'
                "    }\n"
                "  ]\n"
                "}\n\n"
                "TEXTO DE LA SENTENCIA:\n"
                + raw_text[:14000]
            )

            # --- PROCESAMIENTO CON GOOGLE GEMINI ---
            if provider == "Google Gemini (IA Real)":
                clean_key = api_key.strip()
                if not clean_key:
                    st.error("❌ No ingresaste la API Key de Gemini en la barra lateral.")
                elif genai is None:
                    st.error("❌ La librería google-generativeai no está instalada correctamente.")
                else:
                    try:
                        genai.configure(api_key=clean_key)
                        
                        # Detección dinámica de modelos activos compatibles con la clave
                        detected_models = []
                        try:
                            for m in genai.list_models():
                                if 'generateContent' in m.supported_generation_methods:
                                    detected_models.append(m.name)
                        except Exception:
                            pass

                        # Lista priorizada de modelos
                        candidate_list = detected_models + [
                            "gemini-1.5-flash",
                            "models/gemini-1.5-flash",
                            "gemini-1.5-pro",
                            "models/gemini-1.5-pro",
                            "gemini-2.0-flash",
                            "gemini-pro"
                        ]

                        res_text = None
                        last_exception = None

                        for m_name in candidate_list:
                            if not m_name:
                                continue
                            try:
                                try:
                                    model = genai.GenerativeModel(
                                        m_name,
                                        generation_config={"response_mime_type": "application/json"}
                                    )
                                    res = model.generate_content(prompt_text)
                                except Exception:
                                    model = genai.GenerativeModel(m_name)
                                    res = model.generate_content(prompt_text)

                                if res and res.text:
                                    res_text = res.text
                                    model_used_name = m_name
                                    break
                            except Exception as ex:
                                last_exception = ex
                                continue

                        if res_text:
                            clean_json = res_text.replace("```json", "").replace("```", "").strip()
                            data = json.loads(clean_json)
                            criterios_res = data.get("criterios_evaluados", [])
                            hallazgos_res = data.get("hallazgos_parrafos", [])
                            if criterios_res and hallazgos_res:
                                ejecutado_con_ia = True
                        else:
                            st.error(f"❌ No se pudo conectar con Gemini: {last_exception}")

                    except Exception as e:
                        st.error(f"❌ Error de autenticación con la API Key: {e}")

            # --- FALLBACK DEMOSTRACIÓN ---
            if not ejecutado_con_ia:
                if provider != "Google Gemini (IA Real)":
                    st.info("ℹ️ Ejecutando auditoría en Modo Demostración Local.")
                else:
                    st.warning("⚠️ Se mostraron datos de respaldo debido al fallo de conexión.")

                criterios_res = []
                for item in edited_rubrica:
                    max_p = item["Max_Pts"]
                    assigned_p = round(max_p * 0.65, 1)
                    criterios_res.append({
                        "nombre": item["Criterio"],
                        "score": assigned_p,
                        "max_pts": max_p,
                        "nivel": "Observado" if assigned_p < (max_p * 0.75) else "Satisfactorio",
                        "detalle": "Argumentación formal de plantilla; requiere valoración individualizada de los medios probatorios."
                    })
                
                hallazgos_res = [
                    {
                        "criterio_afectado": edited_rubrica[2]["Criterio"] if len(edited_rubrica) > 2 else "Juicio Probatorio",
                        "texto_original": "CONSIDERANDO CUARTO: Que, la prueba documental presentada por la parte demandada no desvirtúa las pretensiones de la demandante...",
                        "vicio": "Motivación Aparente y Defecto de Valoración Probatoria (STC 00728-2008-PHC/TC)",
                        "base_legal": "Art. 197 del Código Procesal Civil Peruano",
                        "propuesta_redaccion": "CONSIDERANDO CUARTO (REVISIÓN SUBSANATORIA): Respecto a la documental N° 04 presentada por la defensa, este despacho advierte que si bien sostiene la falta de responsabilidad, queda desvirtuada al contrastarse de forma conjunta con la prueba testimonial de fojas 45, toda vez que no justifica la omisión del deber de cuidado conforme al Art. 197 del CPC."
                    }
                ]
            else:
                st.success(f"✅ **Análisis en Tiempo Real Exitoso** | Procesado con **Google Gemini** (`{model_used_name}`)")

            # --- CÁLCULO DE MÉTRICAS ---
            total_score = sum(c["score"] for c in criterios_res)
            for c in criterios_res:
                c["pct"] = (c["score"] / c["max_pts"]) * 100 if c["max_pts"] > 0 else 0

            pct_cumplimiento = (total_score / max_score_posible) * 100 if max_score_posible > 0 else 0
            pct_riesgo = round(100.0 - pct_cumplimiento, 1)

            st.markdown("---")
            st.subheader("📊 Métrica Global de Auditoría (Estilo Turnitin)")

            col1, col2, col3 = st.columns(3)
            with col1:
                st.markdown(
                    f"""
                <div class="metric-box">
                    <small>CUMPLIMIENTO NORMATIVO</small>
                    <div class="metric-value">{pct_cumplimiento:.1f}%</div>
                    <small>Score: <b>{total_score:.1f} / {max_score_posible:.1f} pts</b></small>
                </div>
                """,
                    unsafe_allow_html=True,
                )

            with col2:
                st.markdown(
                    f"""
                <div class="metric-box">
                    <small>RIESGO DE NULIDAD / DEFECTOS</small>
                    <div class="metric-risk">{pct_riesgo:.1f}%</div>
                    <small>Párrafos críticos: <b>{len(hallazgos_res)}</b></small>
                </div>
                """,
                    unsafe_allow_html=True,
                )

            with col3:
                nivel_txt = "Sobresaliente / Conforme" if pct_cumplimiento >= 85 else ("Observado / Subsanable" if pct_cumplimiento >= 60 else "Alto Riesgo de Nulidad")
                st.markdown(
                    f"""
                <div class="metric-box">
                    <small>ESTADO DE LA SENTENCIA</small>
                    <div style="font-size: 1.3rem; font-weight:800; color:#1E3A8A; margin-top:15px;">{nivel_txt}</div>
                </div>
                """,
                    unsafe_allow_html=True,
                )

            st.markdown("---")
            st.subheader("📋 Calificación por Criterios (Rúbrica Dinámica Configurada)")
            for c in criterios_res:
                st.write(f"**{c['nombre']}** — `{c['score']} / {c['max_pts']} pts` ({c['pct']:.0f}%)")
                st.progress(min(max(c['pct'] / 100.0, 0.0), 1.0))
                st.caption(f"📌 {c['detalle']}")

            st.markdown("---")
            st.subheader("🔍 Auditoría Párrafo por Párrafo (Subrayado & Opciones de Subsanación)")
            st.info("Extractos detectados con vicios de motivación y sus respectivas propuestas de texto subsanado:")

            for i, h in enumerate(hallazgos_res, 1):
                st.markdown(f"#### 📍 Observación N° {i}: {h['criterio_afectado']}")
                st.markdown(f"<span class='legal-badge'>Vicio: {h['vicio']}</span> | **Base Legal:** {h['base_legal']}", unsafe_allow_html=True)
                
                st.markdown(
                    f"""
                    <div class="turnitin-obs-box">
                        <b>⚠️ PÁRRAFO ORIGINAL OBSERVADO EN LA SENTENCIA:</b><br>
                        <i>"{h['texto_original']}"</i>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                
                st.markdown(
                    f"""
                    <div class="turnitin-fix-box">
                        <b>✅ PROPUESTA DE REDACCIÓN SUBSANATORIA (PARA REEMPLAZAR):</b><br>
                        <b>"{h['propuesta_redaccion']}"</b>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                st.markdown("---")

            # --- BOTÓN DE DESCARGA EN WORD ---
            word_file = create_word_report(
                uploaded_file.name,
                total_score,
                max_score_posible,
                pct_cumplimiento,
                pct_riesgo,
                criterios_res,
                hallazgos_res,
            )

            st.download_button(
                label="📄 Descargar Informe Oficial en Word (.docx)",
                data=word_file,
                file_name=f"Reporte_Turnitin_LexSentencia_{uploaded_file.name}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
            )
