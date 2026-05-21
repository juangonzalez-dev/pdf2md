import streamlit as st
import pymupdf4llm
import tempfile
import os
import zipfile
import io
from pathlib import Path

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="PDF → Markdown",
    page_icon="📄",
    layout="centered",
)

# ── Styles ────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600&display=swap');

html, body, [class*="css"] { font-family: 'IBM Plex Sans', sans-serif; }
.stApp { background: #0d0d0d; color: #e8e8e8; }
header[data-testid="stHeader"] { background: transparent; }

h1 {
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: 1.8rem !important;
    color: #f0f0f0 !important;
    letter-spacing: -0.02em;
    margin-bottom: 0 !important;
}
.subtitle {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.75rem;
    color: #555;
    margin-bottom: 2rem;
    letter-spacing: 0.08em;
}

[data-testid="stFileUploader"] {
    background: #141414;
    border: 1px solid #2a2a2a;
    border-radius: 6px;
    padding: 0.5rem;
}
[data-testid="stFileUploader"]:hover { border-color: #444; }

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

/* Botón primario (Convertir) */
div[data-testid="stButton"] > button[kind="primary"] {
    background: #c8f09a !important;
    color: #0d0d0d !important;
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: 0.85rem !important;
    font-weight: 600 !important;
    letter-spacing: 0.05em;
    border: none !important;
    border-radius: 4px !important;
    padding: 0.6rem 2rem !important;
    width: 100%;
}
div[data-testid="stButton"] > button[kind="primary"]:hover { background: #d8f5b0 !important; }

/* Botón secundario (Limpiar) */
div[data-testid="stButton"] > button[kind="secondary"] {
    background: transparent !important;
    color: #555 !important;
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: 0.8rem !important;
    border: 1px solid #2a2a2a !important;
    border-radius: 4px !important;
    padding: 0.6rem 2rem !important;
    width: 100%;
}
div[data-testid="stButton"] > button[kind="secondary"]:hover {
    border-color: #555 !important;
    color: #ccc !important;
}

/* Botón de descarga */
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
[data-testid="stDownloadButton"] > button:hover { background: #d8f5b0 !important; }

.file-row {
    background: #111;
    border: 1px solid #1e1e1e;
    border-radius: 6px;
    padding: 0.75rem 1rem;
    margin-bottom: 0.5rem;
    display: flex;
    align-items: center;
    justify-content: space-between;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.75rem;
}
.file-name { color: #ccc; }
.file-ok   { color: #c8f09a; }
.file-err  { color: #f09a9a; }
.file-kb   { color: #555; margin-left: 0.5rem; }

.preview-box {
    background: #111;
    border: 1px solid #222;
    border-radius: 6px;
    padding: 1.25rem 1.5rem;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.72rem;
    color: #888;
    line-height: 1.7;
    max-height: 260px;
    overflow-y: auto;
    white-space: pre-wrap;
    word-break: break-word;
}

[data-testid="stExpander"] {
    background: #111 !important;
    border: 1px solid #1e1e1e !important;
    border-radius: 6px !important;
}
hr  { border-color: #1e1e1e !important; }
.stAlert { background: #111 !important; border-radius: 6px !important; }
</style>
""", unsafe_allow_html=True)


# ── Session state ─────────────────────────────────────────────────────────────
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0
if "results" not in st.session_state:
    st.session_state.results = None


# ── Helpers ───────────────────────────────────────────────────────────────────
def estimate_tokens(text: str) -> int:
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

def build_zip(results: list) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, md_text in results:
            zf.writestr(name, md_text.encode("utf-8"))
    return buf.getvalue()

def do_clear():
    st.session_state.uploader_key += 1   # fuerza re-render del uploader
    st.session_state.results = None


# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("# pdf → markdown")
st.markdown('<p class="subtitle">TOKEN-EFFICIENT CONVERSION FOR LLM CONSUMPTION</p>', unsafe_allow_html=True)
st.divider()


# ── Upload ────────────────────────────────────────────────────────────────────
uploaded_files = st.file_uploader(
    "Sube uno o varios PDFs",
    type=["pdf"],
    accept_multiple_files=True,
    label_visibility="collapsed",
    key=f"uploader_{st.session_state.uploader_key}",
)

# ── Botones ───────────────────────────────────────────────────────────────────
st.markdown("")
col_convert, col_clear = st.columns([3, 1])

with col_convert:
    convert_clicked = st.button(
        "⚡ Convertir",
        type="primary",
        disabled=not uploaded_files,
    )

with col_clear:
    st.button(
        "🗑 Limpiar",
        type="secondary",
        on_click=do_clear,
        disabled=not uploaded_files and st.session_state.results is None,
    )


# ── Conversión ────────────────────────────────────────────────────────────────
if convert_clicked and uploaded_files:
    results      = []
    total_pdf_kb = 0
    total_md_kb  = 0
    total_tokens = 0

    progress = st.progress(0, text="Convirtiendo archivos…")

    for i, f in enumerate(uploaded_files):
        pdf_bytes = f.read()
        pdf_kb    = len(pdf_bytes) / 1024
        total_pdf_kb += pdf_kb

        try:
            md_text = pdf_bytes_to_md(pdf_bytes)
            md_kb   = len(md_text.encode("utf-8")) / 1024
            tokens  = estimate_tokens(md_text)
            total_md_kb  += md_kb
            total_tokens += tokens
            stem = Path(f.name).stem
            results.append({
                "md_name":   f"{stem}.md",
                "md_text":   md_text,
                "pdf_kb":    pdf_kb,
                "md_kb":     md_kb,
                "tokens":    tokens,
                "ok":        True,
                "orig_name": f.name,
            })
        except Exception as e:
            results.append({
                "md_name":   f.name,
                "ok":        False,
                "error":     str(e),
                "orig_name": f.name,
                "pdf_kb":    pdf_kb,
            })

        progress.progress((i + 1) / len(uploaded_files),
                          text=f"Convirtiendo {i+1}/{len(uploaded_files)}…")

    progress.empty()
    st.session_state.results = {
        "files":         results,
        "total_pdf_kb":  total_pdf_kb,
        "total_md_kb":   total_md_kb,
        "total_tokens":  total_tokens,
    }


# ── Resultados ────────────────────────────────────────────────────────────────
if st.session_state.results:
    data    = st.session_state.results
    results = data["files"]
    ok_list = [r for r in results if r["ok"]]

    st.divider()

    # Filas de estado
    for r in results:
        if r["ok"]:
            st.markdown(
                f'<div class="file-row">'
                f'<span class="file-name">📄 {r["orig_name"]}</span>'
                f'<span><span class="file-ok">✓ convertido</span>'
                f'<span class="file-kb">{r["pdf_kb"]:.1f} KB → {r["md_kb"]:.1f} KB'
                f' · ~{r["tokens"]:,} tokens</span></span>'
                f'</div>',
                unsafe_allow_html=True,
            )
        else:
            st.markdown(
                f'<div class="file-row">'
                f'<span class="file-name">📄 {r["orig_name"]}</span>'
                f'<span class="file-err">✗ {r["error"]}</span>'
                f'</div>',
                unsafe_allow_html=True,
            )

    # Métricas totales
    if ok_list:
        reduction = (1 - data["total_md_kb"] / data["total_pdf_kb"]) * 100
        st.divider()
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.metric("Archivos", f"{len(ok_list)}/{len(results)}")
        with c2:
            st.metric("PDF total", f"{data['total_pdf_kb']:.1f} KB")
        with c3:
            st.metric("Reducción", f"{reduction:.0f}%",
                      delta=f"-{data['total_pdf_kb'] - data['total_md_kb']:.1f} KB")
        with c4:
            st.metric("~Tokens total", f"{data['total_tokens']:,}")

        st.divider()

        # Descarga
        if len(ok_list) == 1:
            r = ok_list[0]
            st.download_button(
                label="⬇ Descargar .md",
                data=r["md_text"].encode("utf-8"),
                file_name=r["md_name"],
                mime="text/markdown",
            )
            with st.expander("Vista previa", expanded=True):
                preview = r["md_text"][:3000] + ("\n\n…" if len(r["md_text"]) > 3000 else "")
                st.markdown(f'<div class="preview-box">{preview}</div>',
                            unsafe_allow_html=True)
        else:
            zip_bytes = build_zip([(r["md_name"], r["md_text"]) for r in ok_list])
            st.download_button(
                label=f"⬇ Descargar todos ({len(ok_list)} archivos) como .zip",
                data=zip_bytes,
                file_name="markdown_files.zip",
                mime="application/zip",
            )
            for r in ok_list:
                with st.expander(f"Vista previa — {r['md_name']}"):
                    preview = r["md_text"][:2000] + ("\n\n…" if len(r["md_text"]) > 2000 else "")
                    st.markdown(f'<div class="preview-box">{preview}</div>',
                                unsafe_allow_html=True)

elif not uploaded_files:
    st.markdown("")
    st.info("Arrastra uno o varios PDFs arriba, o haz clic para seleccionarlos.", icon="📎")
