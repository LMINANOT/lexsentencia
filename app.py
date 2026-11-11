import io
import json
import docx
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import RGBColor
import pypdf
import streamlit as st

# Importación opcional de motores de IA
try:
    import google.generativeai as genai
except ImportError:
    genai = None

try:
    from openai import OpenAI
except ImportError:
    OpenAI = None

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
    '<div class="sub-title">Auditoría de Motivación Judicial & Cotejo Normativo Peruano<br><b>Análisis Párrafo por Párrafo (Estilo Turnitin) y Rúbricas Personalizables</b></div>',
    unsafe_allow_html=True,
)

# --- CONFIGURACIÓN DE BARRA LATERAL ---
st.sidebar.header("⚙️ Motor de IA & API")
provider = st.sidebar.selectbox("Proveedor de IA", ["Google Gemini", "OpenAI (GPT-4o / GPT-4)", "Modo Demostración Local"])

if provider == "Google Gemini":
    api_key = st.sidebar.text_input("API Key de Google Gemini (Gratis en AI Studio)", type="password")
elif provider == "OpenAI (GPT-4o / GPT-4)":
    api_key = st.sidebar.text_input("API Key de OpenAI (sk-...)", type="password")
else:
    api_key = ""
    st.sidebar.info("💡 El Modo Demostración funciona inmediatamente sin API Key.")

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

    # Resumen Ejecutivo
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

    # Rúbrica
    doc.add_heading("2. EVALUACIÓN DE RÚBRICA PERSONALIZADA", level=1)
    for c in criterios:
        p = doc.add_paragraph()
        p.add_run(f"• {c['nombre']}: ").bold = True
        p.add_run(f"{c['score']} / {c['max_pts']} pts ({c['pct']:.0f}%) — Estado: {c['nivel']}\n")
        p.add_run(f"  Detalle: {c['detalle']}")

    doc.add_paragraph()

    # Auditoría Párrafo por Párrafo
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
        with st.spinner("🔍 Analizando motivación párrafo por párrafo y calculando indicadores..."):
            
            criterios_res = []
            hallazgos_res = []
            ejecutado_con_ia = False

            rubrica_str = "\n".join([f"- {item['Criterio']}: max {item['Max_Pts']} pts" for item in edited_rubrica])

            # Construcción segura del prompt
            prompt_text = (
                "Actúa como Auditor Jurídico de Sentencias Peruanas.\n"
                "Evalúa la sentencia según esta rúbrica:\n"
                + rubrica_str + "\n\n"
                "Formato JSON estricto requerido:\n"
                "{\n"
                '  "criterios_evaluados": [\n'
                '    {"nombre": "nombre", "score": 3.0, "max_pts": 4.0, "nivel": "Aceptable", "detalle": "motivo"}\n'
                "  ],\n"
                '  "hallazgos_parrafos": [\n'
                "    {\n"
                '      "criterio_afectado": "nombre",\n'
                '      "texto_original": "párrafo exacto de la sentencia",\n'
                '      "vicio": "tipo de vicio",\n'
                '      "base_legal": "Art. X",\n'
                '      "propuesta_redaccion": "texto mejorado completo"\n'
                "    }\n"
                "  ]\n"
                "}\n\n"
                "SENTENCIA A AUDITAR:\n"
                + raw_text[:12000]
            )

            # --- PROCESAMIENTO CON GOOGLE GEMINI ---
            if provider == "Google Gemini":
                if not api_key.strip():
                    st.warning("⚠️ No ingresaste una API Key de Gemini. Se usará el Modo Demostración Local.")
                elif genai is None:
                    st.error("La librería google-generativeai no está disponible.")
                else:
                    try:
                        genai.configure(api_key=api_key)
                        candidate_models = ["gemini-1.5-flash", "gemini-2.0-flash", "gemini-1.5-pro"]
                        res_text = None
                        last_err = None
                        
                        for m_name in candidate_models:
                            try:
                                model = genai.GenerativeModel(m_name)
                                res = model.generate_content(prompt_text)
                                if res and res.text:
                                    res_text = res.text
                                    break
                            except Exception as ex:
                                last_err = ex
                                continue

                        if res_text:
                            clean_json = res_text.replace("```json", "").replace("```", "").strip()
                            data = json.loads(clean_json)
                            criterios_res = data["criterios_evaluados"]
                            hallazgos_res = data["hallazgos_parrafos"]
                            ejecutado_con_ia = True
                        else:
                            st.warning(f"⚠️ Error al conectar con Gemini ({last_err}). Usando modo demostración.")
                            
                    except Exception as e:
                        st.warning(f"⚠️ Error con la API Key de Gemini ({e}). Usando modo demostración.")

            # --- PROCESAMIENTO CON OPENAI (GPT-4o) ---
            elif provider == "OpenAI (GPT-4o / GPT-4)":
                if not api_key.strip():
                    st.warning("⚠️ No ingresaste una API Key de OpenAI. Se usará el Modo Demostración Local.")
                elif OpenAI is None:
                    st.error("La librería openai no está instalada.")
                else:
                    try:
                        client = OpenAI(api_key=api_key)
                        response = client.chat.completions.create(
                            model="gpt-4o",
                            messages=[{"role": "user", "content": prompt_text}],
                            response_format={"type": "json_object"}
                        )
                        data = json.loads(response.choices[0].message.content)
                        criterios_res = data["criterios_evaluados"]
                        hallazgos_res = data["hallazgos_parrafos"]
                        ejecutado_con_ia = True
                    except Exception as e:
                        st.warning(f"⚠️ Error con OpenAI ({e}). Usando modo demostración.")

            # --- DEMOSTRACIÓN LOCAL (FALLBACK GARANTIZ
