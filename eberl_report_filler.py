import fitz  # PyMuPDF
import json
import requests
import os
from io import BytesIO
from docx import Document
import streamlit as st

# === SECURE API KEY ===
OPENROUTER_API_KEY = st.secrets.get("OPENROUTER_API_KEY") or os.getenv("OPENROUTER_API_KEY")
if not OPENROUTER_API_KEY:
    st.error("❌ OpenRouter API key not found. Add it in Streamlit → Settings → Secrets.")
    st.stop()

# === PDF TEXT EXTRACTION ===
def extract_pdf_text(uploaded_pdfs):
    combined_text = ""
    for file in uploaded_pdfs:
        with fitz.open(stream=file.read(), filetype="pdf") as doc:
            for page in doc:
                combined_text += page.get_text()
    return combined_text.encode("utf-8", "ignore").decode("utf-8")

# === PLACEHOLDER EXTRACTION FROM DOCX ===
def extract_placeholders(docx_file):
    doc = Document(docx_file)
    placeholders = set()
    for para in doc.paragraphs:
        for word in para.text.split():
            if word.startswith("[") and word.endswith("]"):
                placeholders.add(word.strip("[]"))
    return list(placeholders)

# === CALL LLM TO FILL PLACEHOLDERS ===
def call_llm(pdf_text, placeholders):
    prompt = f"""
You are an insurance report assistant. From the PDF report text below, extract values for the following fields:

{placeholders}

Report:
\"\"\"
{pdf_text[:6000]}
\"\"\"

Return ONLY valid JSON like:
{{
  "XM8_INSURED_NAME": "New Zion Hill Baptist Church",
  "XM8_DATE_INSPECTED": "2024-10-21",
  ...
}}
"""
    headers = {
        "Authorization": f"Bearer {OPENROUTER_API_KEY}",
        "Content-Type": "application/json",
        "HTTP-Referer": "https://your-app-name.streamlit.app"  # Optional
    }

    body = {
        "model": "mistralai/mixtral-8x7b",
        "messages": [{"role": "user", "content": prompt}]
    }

    try:
        encoded_body = json.dumps(body, ensure_ascii=False).encode("utf-8")
        res = requests.post("https://openrouter.ai/api/v1/chat/completions", headers=headers, data=encoded_body)
        res.raise_for_status()
        content = res.json()["choices"][0]["message"]["content"]
        return json.loads(content)
    except Exception as e:
        st.error(f"❌ LLM call failed: {e}")
        return {}

# === MOCK FALLBACK DATA ===
def mock_data():
    return {
        "XM8_INSURED_NAME": "NEW ZION HILL BAPTIST CHURCH",
        "XM8_DATE_LOSS": "2024-10-21",
        "XM8_DATE_INSPECTED": "2024-10-22",
        "XM8_INSURED_P_STREET": "123 Example St.",
        "XM8_INSURED_P_CITY": "Houston",
        "XM8_INSURED_P_STATE": "TX",
        "XM8_INSURED_P_ZIP": "77001",
        "XM8_TOL_DESC": "Wind and Hail",
        "XM8_ESTIMATOR_NAME": "Steven Kujawski",
        "XM8_ESTIMATOR_E_MAIL": "commercialclaims@eberls.com",
        "XM8_ESTIMATOR_C_PHONE": "(800) 607-3604",
        "XM8_DATE_CURRENT": "2024-12-01"
    }

# === FILL DOCX TEMPLATE ===
def fill_template(docx_file, field_values):
    doc = Document(docx_file)
    for para in doc.paragraphs:
        for key, val in field_values.items():
            if f"[{key}]" in para.text:
                para.text = para.text.replace(f"[{key}]", val)
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# === STREAMLIT APP ===
st.set_page_config("Eberl Report Auto-Filler", page_icon="📄")
st.title("📄 Eberl GuideOne Report Auto-Filler")

template_file = st.file_uploader("Upload Template (.docx)", type=["docx"])
pdf_files = st.file_uploader("Upload Photo Reports (.pdf)", type=["pdf"], accept_multiple_files=True)

if st.button("Process"):
    if not template_file or not pdf_files:
        st.error("Please upload both a template and at least one report.")
        st.stop()

    with st.spinner("🔍 Extracting text..."):
        pdf_text = extract_pdf_text(pdf_files)

    with st.spinner("🔎 Finding placeholders..."):
        placeholders = extract_placeholders(template_file)

    with st.spinner("🤖 Calling LLM..."):
        field_values = call_llm(pdf_text, placeholders)
        if not field_values:
            st.warning("⚠️ Using mock data due to LLM issue.")
            field_values = mock_data()

    st.success("✅ Data extracted!")

    with st.spinner("📝 Filling template..."):
        filled_doc = fill_template(template_file, field_values)

    st.download_button(
        label="📥 Download Filled Report",
        data=filled_doc,
        file_name="filled_eberl_report.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
