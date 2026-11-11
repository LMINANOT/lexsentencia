import io
import json
import re
import requests
import docx
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import RGBColor
import pypdf
import streamlit as st

st.set_page_config(
    page_title="LexSentencia 2026 - Auditoría Judicial Peruana",
    page_icon="⚖️",
    layout="wide",
)

# Estilos CSS de la interfaz
st.markdown(
    """
<style>
    .main-title { font-size: 2.2rem; color: #1E3A8A; font-weight: 800; text-align: center; }
    .sub-title { font-size: 1.05rem; color: #4B5563; text-align: center; margin-bottom: 25px; }
    .metric-box { background-color: #F8FAFC; border: 2px solid #1E3A8A; border-radius: 12px; padding: 18px; text-align: center; }
    .metric-value { font-size: 2.8rem; font-weight: 900; color: #1E3A8A; }
    .metric-risk { font-size: 2.8rem; font-weight: 900; color: #DC2626; }
    
    .obs-box {
        background-color: #FEF2F2;
        border-left: 5px solid #EF4444;
        padding: 15px;
        border-radius: 6px;
        margin-top: 15px;
        margin-bottom: 5px;
    }
    .fix-box {
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
    '<div class="sub-title">Auditoría de Motivación Judicial & Cotejo Normativo Peruano<br><b>Análisis Párrafo por Párrafo y Subsanación Puntual de Sentencias Completas</b></div>',
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

    title = doc.add_heading("INFORME DE AUDITORÍA JUDICIAL Y MOTIVACIÓN DE SENTENCIA", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    sub = doc.add_paragraph("LexSentencia 2026 | Estándares del Tribunal Constitucional (STC 00728-2008-PHC/TC)")
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph(f"Documento Auditado: {filename}")
    doc.add_paragraph()

    doc.add_heading("1. RESUMEN EJECUTIVO Y MÉTRICAS DE AUDITORÍA JUDICIAL", level=1)
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

    doc.add_heading("3. AUDITORÍA PÁRRAFO POR PÁRRAFO: OBSERVACIONES Y SUBSANACIÓN JURÍDICA", level=1)
    for i, h in enumerate(hallazgos_parrafos, 1):
        doc.add_heading(f"Observación N° {i}: {h['criterio_afectado']}", level=2)
        
        p_obs = doc.add_paragraph()
        p_obs.add_run("TEXTO ORIGINAL OBSERVADO EN LA SENTENCIA:\n").bold = True
        run_obs = p_obs.add_run(f'"{h["texto_original"]}"')
        run_obs.font.italic = True
        run_obs.font.color.rgb = RGBColor(180, 0, 0)
        
        p_law = doc.add_paragraph()
        p_law.add_run("Vicio Detectado / Base Legal Peruana: ").bold = True
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
    st.success(f"📄 Sentencia cargada: **{uploaded_file.name}** ({len(raw_text):,} caracteres en total)")

    if st.button("🚀 Auditar Sentencia (Análisis de Párrafos y Correcciones Directas)", type="primary"):
        with st.spinner("🔍 Conectando con Google Gemini para auditar la sentencia completa..."):
            
            criterios_res = []
            hallazgos_res = []
            ejecutado_con_ia = False
            model_used_name = ""

            rubrica_str = "\n".join([f"- {item['Criterio']}: max {item['Max_Pts']} pts" for item in edited_rubrica])

            prompt_text = (
                "Actúa como un Auditor Jurídico Experto en Derecho Peruano.\n"
                "Se te adjunta el TEXTO COMPLETO de una sentencia judicial peruana. Realiza un análisis riguroso pero equitativo verificando:\n"
                "- Art. 139 inc. 5 de la Constitución Política del Perú.\n"
                "- Precedente vinculante STC N° 00728-2008-PHC/TC (Vicios de motivación).\n"
                "- Normativa peruana aplicable (CPC, CPP, CC o CP según el caso).\n\n"
                "RÚBRICA DE EVALUACIÓN PERSONALIZADA:\n"
                + rubrica_str + "\n\n"
                "INSTRUCCIONES OBLIGATORIAS DE UBICACIÓN Y SUBSANACIÓN PÁRRAFO POR PÁRRAFO:\n"
                "1. Lee TODO el documento de principio a fin antes de evaluar (no asumas que falta el fallo o la prueba a menos que verdaderamente esté ausente en el texto).\n"
                "2. Identifica entre 3 y 6 hallazgos o párrafos específicos DISTINTOS a lo largo de toda la sentencia donde existan vicios de motivación o deficiencias de fundamentación.\n"
                "3. En 'texto_original', DEBES CITAR EL PÁRRAFO EXACTO Y LITERAL de la sentencia sin alterarlo, anteponiendo la ubicación o número de considerando donde se encuentra. Ejemplo: '[CONSIDERANDO SÉTIMO, PÁRRAFO 2]: Que, con relación a la prueba documental...'\n"
                "4. En 'propuesta_redaccion', redacta la versión corregida y subsanada de ESE MISMO PÁRRAFO O SECCIÓN ESPECÍFICA, de modo que el usuario pueda tomar la propuesta como un reemplazo directo de ese fragmento exacto.\n"
                "5. Responde EXCLUSIVAMENTE un objeto JSON válido con esta estructura exacta (sin texto previo ni marcas ```json):\n"
                "{\n"
                '  "criterios_evaluados": [\n'
                '    {"nombre": "Nombre Criterio", "score": 3.5, "max_pts": 4.0, "nivel": "Conforme", "detalle": "Evaluación objetiva del criterio"}\n'
                "  ],\n"
                '  "hallazgos_parrafos": [\n'
                "    {\n"
                '      "criterio_afectado": "Nombre Criterio",\n'
                '      "texto_original": "[UBICACIÓN EXACTA DEL CONSIDERANDO]: Texto literal e inalterado del párrafo de la sentencia",\n'
                '      "vicio": "Vicio de motivación según STC 00728-2008-PHC/TC",\n'
                '      "base_legal": "Artículos o jurisprudencia peruana aplicable",\n'
                '      "propuesta_redaccion": "Redacción subsanatoria exacta que reemplaza directamente al párrafo observado"\n'
                "    }\n"
                "  ]\n"
                "}\n\n"
                "TEXTO COMPLETO DE LA SENTENCIA A AUDITAR:\n"
                + raw_text[:180000]
            )

            # --- CONEXIÓN DINÁMICA SANITIZADA A GOOGLE GEMINI ---
            if provider == "Google Gemini (IA Real)":
                # Elimina cualquier espacio invisible (\xa0, \t, \r, \n) que altere la URL
                clean_key = re.sub(r'[\s\xa0\ufeff]', '', api_key)
                
                if not clean_key:
                    st.error("❌ No ingresaste la API Key de Gemini en la barra lateral.")
                else:
                    headers = {
                        "Content-Type": "application/json",
                        "x-goog-api-key": clean_key
                    }
                    
                    candidate_models = [
                        ("v1beta", "gemini-1.5-flash"),
                        ("v1beta", "gemini-1.5-pro"),
                        ("v1beta", "gemini-2.0-flash")
                    ]

                    res_text = None
                    last_http_error = ""

                    for api_ver, m_name in candidate_models:
                        # Construcción y desinfección estricta de la URL objetivo
                        raw_url = f"[https://generativelanguage.googleapis.com/](https://generativelanguage.googleapis.com/){api_ver}/models/{m_name}:generateContent"
                        clean_url = re.sub(r'[\s\xa0\ufeff]', '', raw_url)

                        payload_json = {
                            "contents": [{"parts": [{"text": prompt_text}]}],
                            "generationConfig": {
                                "temperature": 0.2
                            }
                        }

                        try:
                            response = requests.post(
                                clean_url,
                                headers=headers,
                                json=payload_json,
                                timeout=60
                            )

                            if response.status_code == 200:
                                res_json = response.json()
                                candidates = res_json.get("candidates", [])
                                if candidates and "content" in candidates[0]:
                                    parts = candidates[0]["content"].get("parts", [])
                                    if parts and "text" in parts[0]:
                                        res_text = parts[0]["text"]
                                        model_used_name = f"{m_name} ({api_ver})"
                                        break
                            else:
                                last_http_error = f"{m_name} [{api_ver}] -> Estado HTTP {response.status_code}: {response.text[:200]}"
                        except Exception as req_e:
                            last_http_error = f"Excepción en {m_name}: {req_e}"

                    if res_text:
                        clean_str = res_text.strip()
                        if "```json" in clean_str:
                            clean_str = clean_str.split("```json")[1].split("```")[0].strip()
                        elif "```" in clean_str:
                            clean_str = clean_str.split("```")[1].split("```")[0].strip()

                        start_idx = clean_str.find("{")
                        end_idx = clean_str.rfind("}")
                        if start_idx != -1 and end_idx != -1:
                            clean_str = clean_str[start_idx:end_idx+1]

                        try:
                            data = json.loads(clean_str)
                            criterios_res = data.get("criterios_evaluados", [])
                            hallazgos_res = data.get("hallazgos_parrafos", [])
                            if criterios_res and hallazgos_res:
                                ejecutado_con_ia = True
                        except Exception as parse_e:
                            st.error(f"❌ La IA respondió pero no se pudo procesar la estructura JSON: {parse_e}")
                    else:
                        st.error(f"❌ Fallo al comunicar con Google API: {last_http_error}")

            # --- MODO DEMOSTRACIÓN LOCAL ---
            if not ejecutado_con_ia:
                if provider != "Google Gemini (IA Real)":
                    st.info("ℹ️ Ejecutando auditoría en Modo Demostración Local.")
                else:
                    st.warning("⚠️ Se activaron datos de prueba debido a un problema con la clave o conexión.")

                criterios_res = []
                for item in edited_rubrica:
                    max_p = item["Max_Pts"]
                    assigned_p = round(max_p * 0.65, 1)
                    criterios_res.append({
                        "nombre": item["Criterio"],
                        "score": assigned_p,
                        "max_pts": max_p,
                        "nivel": "Observado" if assigned_p < (max_p * 0.75) else "Satisfactorio",
                        "detalle": "Presenta argumentación genérica; requiere fundamentación concreta en normas peruanas e interpretación conjunta de pruebas."
                    })
                
                hallazgos_res = [
                    {
                        "criterio_afectado": edited_rubrica[2]["Criterio"] if len(edited_rubrica) > 2 else "Juicio Probatorio",
                        "texto_original": "[CONSIDERANDO SÉTIMO - PÁRRAFO 2]: Que, la prueba documental presentada por la parte demandada no desvirtúa las pretensiones de la demandante...",
                        "vicio": "Motivación Aparente y Defecto de Valoración Probatoria (STC 00728-2008-PHC/TC)",
                        "base_legal": "Art. 197 del Código Procesal Civil Peruano",
                        "propuesta_redaccion": "[REEMPLAZO DIRECTO DEL PÁRRAFO SÉTIMO]: Respecto a la documental N° 04 presentada por la defensa, este despacho advierte que si bien sostiene la falta de responsabilidad, queda desvirtuada al contrastarse de forma conjunta con la prueba testimonial de fojas 45, toda vez que no justifica la omisión del deber de cuidado conforme al Art. 197 del CPC."
                    }
                ]
            else:
                st.success(f"✅ **¡Auditoría Real Exitosa!** Sentencia procesada en tiempo real con **Google Gemini** (`{model_used_name}`)")

            # --- CÁLCULO DE MÉTRICAS ---
            total_score = sum(c["score"] for c in criterios_res)
            for c in criterios_res:
                c["pct"] = (c["score"] / c["max_pts"]) * 100 if c["max_pts"] > 0 else 0

            pct_cumplimiento = (total_score / max_score_posible) * 100 if max_score_posible > 0 else 0
            pct_riesgo = round(100.0 - pct_cumplimiento, 1)

            st.markdown("---")
            st.subheader("📊 Métrica Global de Auditoría Judicial")

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
            st.subheader("🔍 Auditoría Párrafo por Párrafo (Subsanación Integral de Defectos)")
            st.info("Extractos detectados con vicios de motivación y sus respectivas propuestas de texto subsanado conforme a las leyes peruanas:")

            for i, h in enumerate(hallazgos_res, 1):
                st.markdown(f"#### 📍 Observación N° {i}: {h['criterio_afectado']}")
                st.markdown(f"<span class='legal-badge'>Vicio: {h['vicio']}</span> | **Base Legal Peruana:** {h['base_legal']}", unsafe_allow_html=True)
                
                st.markdown(
                    f"""
                    <div class="obs-box">
                        <b>⚠️ PÁRRAFO ORIGINAL OBSERVADO EN LA SENTENCIA:</b><br>
                        <i>"{h['texto_original']}"</i>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                
                st.markdown(
                    f"""
                    <div class="fix-box">
                        <b>✅ PROPUESTA DE REDACCIÓN SUBSANATORIA (CONFORME A DERECHO PERUANO):</b><br>
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
                file_name=f"Informe_Auditoria_LexSentencia_{uploaded_file.name}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
            )
