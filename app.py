from pathlib import Path

code = r'''import io
import json
import re
import time
import requests
import docx
import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import RGBColor

# ============================================================
# LEXSENTENCIA 2026 - AUDITORÍA JUDICIAL POR BLOQUES
# ============================================================

st.set_page_config(
    page_title="LexSentencia 2026 - Auditoría Judicial Peruana",
    page_icon="⚖️",
    layout="wide",
)

st.markdown("""
<style>
.main-title {font-size:2.2rem;color:#1E3A8A;font-weight:800;text-align:center}
.sub-title {font-size:1.05rem;color:#4B5563;text-align:center;margin-bottom:25px}
.metric-box {background:#F8FAFC;border:2px solid #1E3A8A;border-radius:12px;padding:18px;text-align:center}
.metric-value {font-size:2.5rem;font-weight:900;color:#1E3A8A}
.metric-risk {font-size:2.5rem;font-weight:900;color:#DC2626}
.turnitin-obs-box {background:#FEF2F2;border-left:5px solid #EF4444;padding:15px;border-radius:6px;margin:12px 0}
.turnitin-fix-box {background:#ECFDF5;border-left:5px solid #10B981;padding:15px;border-radius:6px;margin:12px 0}
.legal-badge {background:#DBEAFE;color:#1E40AF;padding:3px 8px;border-radius:4px;font-weight:700;font-size:.82rem}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">⚖️ LexSentencia 2026</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="sub-title">Auditoría de Motivación Judicial & Cotejo Normativo Peruano'
    '<br><b>Análisis integral por bloques con Google Gemini API</b></div>',
    unsafe_allow_html=True,
)

# ---------------- CONFIGURACIÓN ----------------
st.sidebar.header("⚙️ Motor de IA & API")
provider = st.sidebar.selectbox(
    "Proveedor de IA",
    ["Google Gemini (IA Real)", "Modo Demostración Local"]
)

api_key = ""
if provider == "Google Gemini (IA Real)":
    api_key = st.sidebar.text_input(
        "API Key de Google Gemini",
        type="password",
        value="",
        help="También puedes guardarla en Streamlit Secrets como GEMINI_API_KEY."
    )
else:
    st.sidebar.info("El modo demostración genera un ejemplo ficticio. No constituye auditoría jurídica.")

st.sidebar.markdown("---")
st.sidebar.header("🛠️ Configuración de Rúbrica")

DEFAULT_RUBRIC = [
    {"Criterio": "1. Delimitación del problema jurídico", "Max_Pts": 4.0},
    {"Criterio": "2. Coherencia lógica y congruencia procesal", "Max_Pts": 4.0},
    {"Criterio": "3. Razonamiento probatorio y valoración", "Max_Pts": 4.0},
    {"Criterio": "4. Solidez argumentativa interna y externa", "Max_Pts": 4.0},
    {"Criterio": "5. Utilización pertinente de jurisprudencia peruana", "Max_Pts": 4.0},
]
if "rubrica_criterios" not in st.session_state:
    st.session_state.rubrica_criterios = DEFAULT_RUBRIC

edited_rubrica = st.sidebar.data_editor(
    st.session_state.rubrica_criterios,
    num_rows="dynamic",
    column_config={
        "Criterio": st.column_config.TextColumn("Nombre del criterio", width="large"),
        "Max_Pts": st.column_config.NumberColumn(
            "Puntos máximos", min_value=0.0, max_value=20.0, step=0.5
        ),
    },
    key="rubrica_editor",
)
edited_rubrica = [
    x for x in edited_rubrica
    if str(x.get("Criterio", "")).strip() and float(x.get("Max_Pts") or 0) > 0
]
max_score_posible = sum(float(x["Max_Pts"]) for x in edited_rubrica)
st.sidebar.caption(f"Puntaje máximo configurado: {max_score_posible:.1f} puntos")

# El tamaño de bloque es un parámetro técnico, no un límite total del documento.
chunk_size = st.sidebar.select_slider(
    "Tamaño aproximado de cada bloque (caracteres)",
    options=[8000, 10000, 12000, 14000],
    value=10000,
)
overlap_chars = st.sidebar.select_slider(
    "Solapamiento entre bloques",
    options=[400, 600, 800, 1000],
    value=800,
)

# ---------------- EXTRACCIÓN Y DIVISIÓN ----------------
def extract_text(uploaded_file) -> str:
    name = uploaded_file.name.lower()
    if name.endswith(".pdf"):
        reader = pypdf.PdfReader(uploaded_file)
        pages = []
        for page_number, page in enumerate(reader.pages, start=1):
            page_text = page.extract_text() or ""
            pages.append(f"\n[PÁGINA PDF {page_number}]\n{page_text}")
        return "\n".join(pages)
    if name.endswith(".docx"):
        document = docx.Document(uploaded_file)
        return "\n".join(p.text for p in document.paragraphs if p.text.strip())
    return uploaded_file.read().decode("utf-8", errors="ignore")


def split_into_chunks(text: str, max_chars: int, overlap: int):
    """Divide por párrafos; conserva referencias consecutivas y solapamiento."""
    paragraphs = [
        p.strip() for p in re.split(r"\n\s*\n", text)
        if p and p.strip()
    ]
    if not paragraphs:
        paragraphs = [text.strip()] if text.strip() else []

    tagged = [f"[PÁRRAFO {i:05d}]\n{p}" for i, p in enumerate(paragraphs, 1)]
    chunks = []
    current = []
    current_len = 0

    for paragraph in tagged:
        p_len = len(paragraph) + 2
        if current and current_len + p_len > max_chars:
            chunks.append("\n\n".join(current))

            # Retener párrafos completos del final como contexto compartido.
            tail = []
            tail_len = 0
            for old in reversed(current):
                if tail_len + len(old) + 2 > overlap:
                    break
                tail.insert(0, old)
                tail_len += len(old) + 2

            current = tail[:]
            current_len = sum(len(x) + 2 for x in current)

        # Un párrafo individual muy largo se conserva completo, sin truncarlo.
        current.append(paragraph)
        current_len += p_len

    if current:
        chunks.append("\n\n".join(current))
    return chunks


def normalize_text(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


# ---------------- GEMINI ----------------
GEMINI_BASE = "https://generativelanguage.googleapis.com/v1beta"


def list_usable_models(key: str):
    response = requests.get(
        f"{GEMINI_BASE}/models",
        params={"key": key, "pageSize": 1000},
        timeout=30,
    )
    response.raise_for_status()
    models = response.json().get("models", [])
    usable = [
        m for m in models
        if "generateContent" in m.get("supportedGenerationMethods", [])
        and "flash" in m.get("name", "").lower()
    ]
    if not usable:
        usable = [
            m for m in models
            if "generateContent" in m.get("supportedGenerationMethods", [])
        ]

    # Preferir versiones estables recientes cuando la clave las tenga habilitadas.
    preferred = [
        "gemini-3.8-flash",
        "gemini-3.7-flash",
        "gemini-3.6-flash",
        "gemini-3.5-flash",
        "gemini-2.5-flash",
    ]
    by_name = {m["name"].replace("models/", ""): m for m in usable}
    for name in preferred:
        if name in by_name:
            return name, by_name[name]

    if not usable:
        raise RuntimeError("La API Key no muestra modelos compatibles con generateContent.")

    usable.sort(key=lambda x: x.get("name", ""), reverse=True)
    chosen = usable[0]
    return chosen["name"].replace("models/", ""), chosen


def call_gemini(model_name: str, key: str, prompt: str, max_output_tokens=8192):
    url = f"{GEMINI_BASE}/models/{model_name}:generateContent"
    payload = {
        "contents": [{"role": "user", "parts": [{"text": prompt}]}],
        "generationConfig": {
            "temperature": 0.1,
            "responseMimeType": "application/json",
            "maxOutputTokens": max_output_tokens,
        },
    }

    last_error = ""
    for attempt in range(3):
        try:
            response = requests.post(
                url,
                params={"key": key},
                headers={"Content-Type": "application/json"},
                json=payload,
                timeout=240,
            )
            if response.status_code in (429, 500, 502, 503, 504):
                last_error = f"HTTP {response.status_code}: {response.text[:1200]}"
                time.sleep(2 ** attempt)
                continue
            response.raise_for_status()
            result = response.json()
            parts = result["candidates"][0]["content"]["parts"]
            text = "".join(part.get("text", "") for part in parts).strip()
            if text.startswith("```"):
                text = re.sub(r"^```(?:json)?\s*", "", text)
                text = re.sub(r"\s*```$", "", text)
            return json.loads(text)
        except (requests.RequestException, KeyError, IndexError, ValueError) as exc:
            last_error = str(exc)
            if attempt < 2:
                time.sleep(2 ** attempt)

    raise RuntimeError(last_error or "No se obtuvo respuesta válida de Gemini.")


def audit_chunk(model_name: str, key: str, chunk_text: str, chunk_number: int, total_chunks: int):
    prompt = f"""
Actúa como auditor jurídico experto en derecho peruano.
Analiza exclusivamente el contenido del bloque {chunk_number} de {total_chunks}
de una sentencia. Este bloque forma parte de un documento mayor: NO concluyas
sobre la sentencia completa ni asignes una nota global.

Examina posibles defectos de motivación, congruencia, valoración probatoria,
coherencia argumentativa y uso de normas/jurisprudencia. Distingue una verdadera
deficiencia de una mera discrepancia con el criterio judicial. No inventes hechos,
citas, artículos, precedentes ni extractos. Si no hay un defecto sustentable,
devuelve una lista vacía.

Devuelve únicamente JSON con esta estructura:
{{
  "hallazgos_parrafos": [
    {{
      "parrafo_ref": "PÁRRAFO 00001",
      "criterio_afectado": "criterio",
      "texto_original": "extracto literal breve y verificable",
      "vicio": "defecto identificado o 'No se acredita un vicio'",
      "base_legal": "norma o jurisprudencia pertinente; indica si requiere verificación",
      "analisis": "explicación jurídica contextual",
      "propuesta_redaccion": "propuesta de subsanación; si no corresponde, cadena vacía"
    }}
  ],
  "observaciones_generales_bloque": ["observación breve"]
}}

Reglas:
- Máximo 8 hallazgos relevantes por bloque.
- El texto_original debe ser literal y estar dentro del bloque.
- No conviertas automáticamente cada párrafo en un hallazgo.
- No atribuyas carácter vinculante a jurisprudencia sin sustento.
- Si falta contexto de otro bloque, dilo expresamente.

CONTENIDO DEL BLOQUE:
{chunk_text}
"""
    return call_gemini(model_name, key, prompt, 8192)


def consolidate_audit(model_name: str, key: str, rubric, findings, document_chars: int):
    rubric_text = "\n".join(
        f"- {item['Criterio']}: máximo {float(item['Max_Pts']):.1f} puntos"
        for item in rubric
    )
    findings_json = json.dumps(findings, ensure_ascii=False)
    prompt = f"""
Actúa como auditor jurídico experto en derecho peruano. Recibirás hallazgos
preliminares de todos los bloques de una sentencia de {document_chars} caracteres.
Realiza una evaluación GLOBAL, no sumes notas parciales de bloques.

RÚBRICA:
{rubric_text}

HALLAZGOS PRELIMINARES:
{findings_json}

INSTRUCCIONES:
1. Evalúa cada criterio de la rúbrica una sola vez, sobre el documento completo.
2. Asigna un puntaje entre 0 y el máximo del criterio, sin excederlo.
3. Fundamenta cada nota con evidencia concreta y explica también fortalezas.
4. Consolida hallazgos repetidos y evita duplicados.
5. No inventes normas, citas, hechos ni pasajes de la sentencia.
6. Si la base legal requiere comprobación, indícalo.
7. El porcentaje de riesgo no equivale automáticamente a probabilidad jurídica de nulidad.
8. Si no hay hallazgos, la lista puede quedar vacía; eso no impide una auditoría válida.

Devuelve únicamente JSON:
{{
  "criterios_evaluados": [
    {{
      "nombre": "nombre exacto del criterio",
      "score": 0.0,
      "max_pts": 4.0,
      "nivel": "Satisfactorio / Observado / Deficiente",
      "detalle": "fundamentación global"
    }}
  ],
  "hallazgos_parrafos": [
    {{
      "parrafo_ref": "PÁRRAFO 00001",
      "criterio_afectado": "criterio",
      "texto_original": "extracto literal",
      "vicio": "defecto",
      "base_legal": "sustento",
      "analisis": "explicación",
      "propuesta_redaccion": "propuesta"
    }}
  ],
  "conclusion_general": "conclusión jurídica global y prudente",
  "limitaciones": ["limitación si corresponde"]
}}
"""
    return call_gemini(model_name, key, prompt, 12000)


# ---------------- INFORME WORD ----------------
def create_word_report(
    filename, score_total, max_score, pct_compliance, pct_risk,
    criteria, findings, conclusion, limitations, model_name, total_chars, total_chunks
):
    doc = Document()
    title = doc.add_heading("INFORME DE AUDITORÍA JUDICIAL Y MOTIVACIÓN DE SENTENCIA", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    subtitle = doc.add_paragraph(
        "LexSentencia 2026 | Referencia de análisis: Constitución Política del Perú "
        "y jurisprudencia constitucional aplicable"
    )
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph(f"Documento auditado: {filename}")
    doc.add_paragraph(f"Extensión extraída: {total_chars:,} caracteres")
    doc.add_paragraph(f"Bloques procesados: {total_chunks}")
    doc.add_paragraph(f"Modelo Gemini: {model_name}")
    doc.add_paragraph(
        "Nota metodológica: la herramienta es un apoyo de análisis y no sustituye "
        "la revisión profesional ni la verificación de fuentes jurídicas."
    )

    doc.add_heading("1. Resumen ejecutivo", level=1)
    table = doc.add_table(rows=3, cols=2)
    table.style = "Table Grid"
    table.cell(0, 0).text = "Métrica"
    table.cell(0, 1).text = "Resultado"
    table.cell(1, 0).text = "Cumplimiento según rúbrica"
    table.cell(1, 1).text = f"{pct_compliance:.1f}% ({score_total:.1f}/{max_score:.1f})"
    table.cell(2, 0).text = "Indicador complementario de aspectos observados"
    table.cell(2, 1).text = f"{pct_risk:.1f}%"

    doc.add_heading("2. Calificación por criterios", level=1)
    for criterion in criteria:
        p = doc.add_paragraph()
        p.add_run(
            f"{criterion['nombre']}: {criterion['score']} / "
            f"{criterion['max_pts']} puntos ({criterion['pct']:.1f}%)"
        ).bold = True
        doc.add_paragraph(
            f"Nivel: {criterion.get('nivel', 'No especificado')}. "
            f"Fundamento: {criterion.get('detalle', '')}"
        )

    doc.add_heading("3. Hallazgos por párrafo", level=1)
    if not findings:
        doc.add_paragraph(
            "No se identificaron hallazgos que superaran el umbral de relevancia "
            "en el análisis automatizado. Esto no equivale a certificar la ausencia "
            "de errores jurídicos."
        )

    for index, finding in enumerate(findings, 1):
        doc.add_heading(
            f"Observación {index}: {finding.get('criterio_afectado', 'Sin criterio')}",
            level=2
        )
        doc.add_paragraph(f"Referencia: {finding.get('parrafo_ref', 'No identificada')}")
        p = doc.add_paragraph()
        p.add_run("Extracto literal: ").bold = True
        run = p.add_run(finding.get("texto_original", ""))
        run.italic = True
        run.font.color.rgb = RGBColor(180, 0, 0)

        p = doc.add_paragraph()
        p.add_run("Vicio / análisis: ").bold = True
        p.add_run(
            f"{finding.get('vicio', '')}. "
            f"{finding.get('analisis', '')}"
        )

        p = doc.add_paragraph()
        p.add_run("Base legal: ").bold = True
        p.add_run(finding.get("base_legal", ""))

        if finding.get("propuesta_redaccion"):
            p = doc.add_paragraph()
            p.add_run("Propuesta de subsanación: ").bold = True
            run = p.add_run(finding["propuesta_redaccion"])
            run.bold = True
            run.font.color.rgb = RGBColor(0, 100, 0)

    doc.add_heading("4. Conclusión general", level=1)
    doc.add_paragraph(conclusion or "No se generó una conclusión general.")

    if limitations:
        doc.add_heading("5. Limitaciones y verificaciones pendientes", level=1)
        for item in limitations:
            doc.add_paragraph(item, style="List Bullet")

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


# ---------------- CARGA Y EJECUCIÓN ----------------
uploaded_file = st.file_uploader(
    "📁 Cargar sentencia judicial (PDF, Word o TXT)",
    type=["pdf", "docx", "txt"]
)

if uploaded_file:
    raw_text = extract_text(uploaded_file)
    total_chars = len(raw_text)
    st.success(
        f"Sentencia cargada: **{uploaded_file.name}** "
        f"({total_chars:,} caracteres extraídos)."
    )

    if total_chars == 0:
        st.error("No se pudo extraer texto. Si el PDF es escaneado, necesita OCR.")
        st.stop()

    if total_chars < 200000:
        st.info(
            f"El documento tiene {total_chars:,} caracteres. "
            "No existe un recorte fijo de 14,000 caracteres: se procesará "
            "todo el texto extraído, aunque sea menor o mayor a 200,000."
        )
    else:
        st.success("El documento supera los 200,000 caracteres y se procesará completo por bloques.")

    if st.button("🚀 Auditar sentencia completa", type="primary"):
        if not edited_rubrica or max_score_posible <= 0:
            st.error("Configura al menos un criterio válido en la rúbrica.")
            st.stop()

        if provider == "Google Gemini (IA Real)":
            # Prioridad a Streamlit Secrets; si no existe, usa el campo lateral.
            secret_key = ""
            try:
                secret_key = st.secrets.get("GEMINI_API_KEY", "")
            except Exception:
                secret_key = ""
            effective_key = secret_key or api_key.strip()

            if not effective_key:
                st.error("Introduce tu API Key o configúrala en Streamlit Secrets.")
                st.stop()

            try:
                with st.spinner("Consultando modelos disponibles de Gemini..."):
                    model_name, model_info = list_usable_models(effective_key)

                st.info(
                    f"Modelo seleccionado automáticamente: **{model_name}**. "
                    "Se verificó que admite generateContent."
                )

                chunks = split_into_chunks(raw_text, chunk_size, overlap_chars)
                st.write(
                    f"El documento se dividirá en **{len(chunks)} bloques**. "
                    "Todos los bloques se enviarán para análisis."
                )

                findings_all = []
                progress = st.progress(0)
                status = st.empty()

                for idx, chunk in enumerate(chunks, start=1):
                    status.write(f"Analizando bloque {idx} de {len(chunks)}...")
                    chunk_result = audit_chunk(
                        model_name, effective_key, chunk, idx, len(chunks)
                    )
                    for finding in chunk_result.get("hallazgos_parrafos", []):
                        quote = normalize_text(finding.get("texto_original", ""))
                        # Validar que el extracto exista en el texto extraído.
                        if quote and quote in normalize_text(raw_text):
                            findings_all.append(finding)
                    progress.progress(idx / len(chunks))

                status.write("Consolidando hallazgos y calificando la sentencia completa...")
                final_data = consolidate_audit(
                    model_name, effective_key, edited_rubrica,
                    findings_all, total_chars
                )

                criteria = final_data.get("criterios_evaluados", [])
                findings = final_data.get("hallazgos_parrafos", [])
                conclusion = final_data.get("conclusion_general", "")
                limitations = final_data.get("limitaciones", [])

                # Verificación estructural de los criterios devueltos.
                expected_names = [x["Criterio"] for x in edited_rubrica]
                returned_names = [x.get("nombre") for x in criteria]
                if not criteria or any(name not in returned_names for name in expected_names):
                    st.error(
                        "Gemini no devolvió todos los criterios de la rúbrica. "
                        "No se mostrará una nota incompleta como resultado final."
                    )
                    st.stop()

                # Normalizar máximos y puntajes conforme a la rúbrica configurada.
                rubric_map = {x["Criterio"]: float(x["Max_Pts"]) for x in edited_rubrica}
                normalized_criteria = []
                for item in criteria:
                    name = item["nombre"]
                    if name not in rubric_map:
                        continue
                    maximum = rubric_map[name]
                    try:
                        score = float(item.get("score", 0))
                    except (TypeError, ValueError):
                        score = 0.0
                    score = min(max(score, 0.0), maximum)
                    normalized_criteria.append({
                        "nombre": name,
                        "score": round(score, 2),
                        "max_pts": maximum,
                        "pct": (score / maximum * 100) if maximum else 0,
                        "nivel": item.get("nivel", "No especificado"),
                        "detalle": item.get("detalle", ""),
                    })

                criteria = normalized_criteria
                total_score = sum(x["score"] for x in criteria)
                pct_compliance = (
                    total_score / max_score_posible * 100
                    if max_score_posible else 0
                )
                # Indicador matemático complementario, no probabilidad de nulidad.
                pct_risk = round(100 - pct_compliance, 1)

                st.success(
                    f"Auditoría real completada. Modelo: {model_name}. "
                    f"Bloques procesados: {len(chunks)}/{len(chunks)}."
                )

            except Exception as exc:
                st.error(
                    f"No se completó la auditoría real: {exc}. "
                    "No se generarán notas ficticias."
                )
                st.stop()

        else:
            st.warning(
                "MODO DEMOSTRACIÓN: los siguientes datos son ilustrativos, "
                "no proceden de una auditoría real de la sentencia."
            )
            criteria = []
            for item in edited_rubrica:
                maximum = float(item["Max_Pts"])
                score = round(maximum * 0.65, 1)
                criteria.append({
                    "nombre": item["Criterio"],
                    "score": score,
                    "max_pts": maximum,
                    "pct": score / maximum * 100 if maximum else 0,
                    "nivel": "Demostración",
                    "detalle": "Resultado ficticio de demostración; no usar como evaluación jurídica.",
                })
            findings = []
            conclusion = "Resultado de demostración local. No es una auditoría real."
            limitations = ["Los puntajes y textos de este modo son ficticios."]
            model_name = "Modo demostración"
            total_score = sum(x["score"] for x in criteria)
            pct_compliance = total_score / max_score_posible * 100 if max_score_posible else 0
            pct_risk = round(100 - pct_compliance, 1)
            chunks = split_into_chunks(raw_text, chunk_size, overlap_chars)

        # ---------------- RESULTADOS ----------------
        st.markdown("---")
        st.subheader("📊 Métrica global de auditoría")
        col1, col2, col3 = st.columns(3)

        with col1:
            st.markdown(
                f'<div class="metric-box"><small>CUMPLIMIENTO SEGÚN RÚBRICA</small>'
                f'<div class="metric-value">{pct_compliance:.1f}%</div>'
                f'<small><b>{total_score:.1f} / {max_score_posible:.1f} puntos</b></small></div>',
                unsafe_allow_html=True,
            )
        with col2:
            st.markdown(
                f'<div class="metric-box"><small>INDICADOR COMPLEMENTARIO</small>'
                f'<div class="metric-risk">{pct_risk:.1f}%</div>'
                f'<small>Hallazgos: <b>{len(findings)}</b></small></div>',
                unsafe_allow_html=True,
            )
        with col3:
            status_text = (
                "Conforme según rúbrica" if pct_compliance >= 85
                else "Requiere revisión" if pct_compliance >= 60
                else "Deficiencias relevantes"
            )
            st.markdown(
                f'<div class="metric-box"><small>RESULTADO DE LA RÚBRICA</small>'
                f'<div style="font-size:1.25rem;font-weight:800;color:#1E3A8A;margin-top:15px">'
                f'{status_text}</div></div>',
                unsafe_allow_html=True,
            )

        st.caption(
            "El indicador complementario se calcula como 100% menos el cumplimiento "
            "de la rúbrica. No representa una probabilidad estadística ni una predicción "
            "de nulidad judicial."
        )

        st.markdown("---")
        st.subheader("📋 Calificación por criterios")
        for item in criteria:
            st.write(
                f"**{item['nombre']}** — {item['score']} / "
                f"{item['max_pts']} puntos ({item['pct']:.1f}%)"
            )
            st.progress(min(max(item["pct"] / 100, 0), 1))
            st.caption(item.get("detalle", ""))

        st.markdown("---")
        st.subheader("🔍 Hallazgos de auditoría")
        if not findings:
            st.info(
                "No se reportaron hallazgos sustentados en los bloques analizados. "
                "Esto no certifica que la sentencia esté libre de defectos."
            )

        for i, finding in enumerate(findings, 1):
            st.markdown(
                f"#### Observación {i}: {finding.get('criterio_afectado', 'Sin criterio')}"
            )
            st.caption(f"Referencia: {finding.get('parrafo_ref', 'No identificada')}")
            st.markdown(
                f"<div class='turnitin-obs-box'><b>Extracto original:</b><br>"
                f"<i>{finding.get('texto_original', '')}</i></div>",
                unsafe_allow_html=True,
            )
            st.markdown(
                f"<span class='legal-badge'>Vicio: {finding.get('vicio', '')}</span> "
                f"**Base legal:** {finding.get('base_legal', '')}",
                unsafe_allow_html=True,
            )
            st.write(f"**Análisis:** {finding.get('analisis', '')}")
            if finding.get("propuesta_redaccion"):
                st.markdown(
                    f"<div class='turnitin-fix-box'><b>Propuesta de subsanación:</b><br>"
                    f"{finding.get('propuesta_redaccion', '')}</div>",
                    unsafe_allow_html=True,
                )
            st.markdown("---")

        st.subheader("⚖️ Conclusión general")
        st.write(conclusion)

        if limitations:
            with st.expander("Limitaciones y verificaciones pendientes"):
                for limitation in limitations:
                    st.write(f"- {limitation}")

        word_file = create_word_report(
            uploaded_file.name,
            total_score,
            max_score_posible,
            pct_compliance,
            pct_risk,
            criteria,
            findings,
            conclusion,
            limitations,
            model_name,
            total_chars,
            len(chunks),
        )

        st.download_button(
            label="📄 Descargar informe de auditoría en Word",
            data=word_file,
            file_name=f"LexSentencia_Auditoria_{uploaded_file.name}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
        )
'''
path = Path("/mnt/data/app_lexsentencia_mejorado.py")
path.write_text(code, encoding="utf-8")
print(f"Archivo creado: {path} ({path.stat().st_size:,} bytes)")
