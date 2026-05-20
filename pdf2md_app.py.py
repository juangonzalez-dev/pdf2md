import streamlit as st
import pymupdf4llm
import tempfile
import os
from pathlib import Path

# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="PDF → Markdown",
    page_icon="📄",
    layout="centered",
)

# ── Styles ────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600&display=swap');

html, body, [class*="css"] {
    font-family: 'IBM Plex Sans', sans-serif;
}

/* Background */
.stApp {
    background: #0d0d0d;
    color: #e8e8e8;
}

/* Hide default header */
header[data-testid="stHeader"] { background: transparent; }

/* Title */
h1 { 
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: 1.8rem !important;
    color: #f0f0f0 !important;
    letter-spacing: -0.02em;
    margin-bottom: 0 !important;
}

/* Subtext */
.subtitle {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.75rem;
    color: #555;
    margin-bottom: 2rem;
    letter-spacing: 0.08em;
}

/* Upload box */
[data-testid="stFileUploader"] {
    background: #141414;
    border: 1px solid #2a2a2a;
    border-radius: 6px;
    padding: 0.5rem;
}
[data-testid="stFileUploader"]:hover {
    border-color: #444;
}

/* Metrics */
[data-testid="stMetric"] {
    background: #141414;
    border: 1px solid #222;
    border-radius: 6px;
    padding: 0.75rem 1rem;
}
[data-testid="stMetricLabel"] {
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: 0.65rem !important;
    color: #555 !important;
    text-transform: uppercase;
    letter-spacing: 0.1em;
}
[data-testid="stMetricValue"] {
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: 1.1rem !important;
    color: #c8f09a !important;
}
[data-testid="stMetricDelta"] {
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: 0.7rem !important;
}

/* Download button */
[data-testid="stDownloadButton"] > button {
    background: #c8f09a !important;
    color: #0d0d0d !important;
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: 0.8rem !important;
    font-weight: 600 !important;
    letter-spacing: 0.05em;
    border: none !important;
    border-radius: 4px !important;
    padding: 0.5rem 1.5rem !important;
    width: 100%;
}
[data-testid="stDownloadButton"] > button:hover {
    background: #d8f5b0 !important;
}

/* Preview area */
.preview-box {
    background: #111;
    border: 1px solid #222;
    border-radius: 6px;
    padding: 1.25rem 1.5rem;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.72rem;
    color: #888;
    line-height: 1.7;
    max-height: 320px;
    overflow-y: auto;
    white-space: pre-wrap;
    word-break: break-word;
}

/* Expander */
[data-testid="stExpander"] {
    background: #111 !important;
    border: 1px solid #1e1e1e !important;
    border-radius: 6px !important;
}

/* Divider */
hr { border-color: #1e1e1e !important; }

/* Success / info */
.stAlert {
    background: #111 !important;
    border-radius: 6px !important;
}
</style>
""", unsafe_allow_html=True)


# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("# pdf → markdown")
st.markdown('<p class="subtitle">TOKEN-EFFICIENT CONVERSION FOR LLM CONSUMPTION</p>', unsafe_allow_html=True)
st.divider()


# ── Upload ────────────────────────────────────────────────────────────────────
uploaded = st.file_uploader(
    "Sube un PDF",
    type=["pdf"],
    label_visibility="collapsed",
)


# ── Conversion logic ──────────────────────────────────────────────────────────
def estimate_tokens(text: str) -> int:
    """Rough estimate: ~4 chars per token."""
    return len(text) // 4

def pdf_bytes_to_md(pdf_bytes: bytes) -> str:
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
        tmp.write(pdf_bytes)
        tmp_path = tmp.name
    try:
        md = pymupdf4llm.to_markdown(tmp_path)
    finally:
        os.unlink(tmp_path)
    return md


if uploaded is not None:
    pdf_bytes = uploaded.read()
    pdf_size_kb = len(pdf_bytes) / 1024

    with st.spinner("Convirtiendo..."):
        try:
            md_text = pdf_bytes_to_md(pdf_bytes)
        except Exception as e:
            st.error(f"Error al convertir: {e}")
            st.stop()

    md_bytes = md_text.encode("utf-8")
    md_size_kb = len(md_bytes) / 1024
    reduction = (1 - md_size_kb / pdf_size_kb) * 100
    tokens_est = estimate_tokens(md_text)
    char_count = len(md_text)

    st.divider()

    # Metrics
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.metric("PDF original", f"{pdf_size_kb:.1f} KB")
    with c2:
        st.metric("Markdown", f"{md_size_kb:.1f} KB")
    with c3:
        st.metric("Reducción", f"{reduction:.0f}%", delta=f"-{pdf_size_kb - md_size_kb:.1f} KB")
    with c4:
        st.metric("~Tokens", f"{tokens_est:,}")

    st.divider()

    # Download
    stem = Path(uploaded.name).stem
    st.download_button(
        label="⬇ Descargar .md",
        data=md_bytes,
        file_name=f"{stem}.md",
        mime="text/markdown",
    )

    # Preview
    st.markdown("")
    with st.expander("Vista previa (primeras 3,000 caracteres)", expanded=True):
        preview = md_text[:3000] + ("\n\n…" if len(md_text) > 3000 else "")
        st.markdown(f'<div class="preview-box">{preview}</div>', unsafe_allow_html=True)

else:
    st.markdown("")
    st.info("Arrastra un PDF arriba o haz clic para seleccionarlo.", icon="📎")
