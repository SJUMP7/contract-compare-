"""
app.py — Compare Hotel Contracts  (optimised build)
Streaming AI response · Cached PDF · Cached model · Clean light UI
"""
import os, re, json, toml
import streamlit as st
from utils import extract_pdf_text, stream_contract_comparison, validate_api_key, detect_available_model
from excel_generator import generate_comparison_excel

# ─── Page config ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Contract Compare",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── CSS ─────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
html,body,[class*="css"]{font-family:'Inter',-apple-system,BlinkMacSystemFont,sans-serif!important}

/* Backgrounds */
.stApp,.block-container{background:#f0f4f8!important}
.block-container{padding:2rem 2.5rem 3rem!important}
section[data-testid="stSidebar"]{background:#ffffff!important;border-right:1px solid #e2e8f0!important}

/* Hero */
.hero{text-align:center;padding:44px 16px 28px}
.eyebrow{font-size:11px;font-weight:700;letter-spacing:.12em;text-transform:uppercase;color:#3b82f6;margin-bottom:8px}
.h1{
    font-size:48px;font-weight:800;letter-spacing:-.025em;line-height:1.1;margin-bottom:12px;
    background: linear-gradient(135deg, #0ea5e9, #6366f1, #a855f7);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    display: inline-block;
}
.sub{font-size:16px;color:#475569;max-width:500px;margin:0 auto;line-height:1.65}

/* Tags */
.tags{display:flex;gap:7px;flex-wrap:wrap;justify-content:center;margin:20px 0 28px}
.tag{font-size:12px;font-weight:500;color:#334155;background:#e0f2fe;border:1px solid #bae6fd;padding:4px 12px;border-radius:100px}

/* Upload card */
.card{background:#ffffff;border:1px solid #e2e8f0;border-radius:16px;padding:22px 20px 18px;transition:all .3s ease;box-shadow: 0 4px 6px -1px rgba(0,0,0,.05)}
.card:hover{border-color:#3b82f6;box-shadow:0 10px 15px -3px rgba(59,130,246,.15), 0 4px 6px -4px rgba(59,130,246,.1)}
.c-eye{font-size:10px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:#3b82f6;margin-bottom:5px}
.c-ttl{font-size:19px;font-weight:600;color:#0f172a;margin-bottom:14px}

/* Buttons */
button[data-testid="baseButton-primary"]{
    background: linear-gradient(135deg, #0ea5e9, #3b82f6)!important;color:#fff!important;border:none!important;
    border-radius:12px!important;font-size:15px!important;font-weight:600!important;
    box-shadow:0 4px 14px rgba(59,130,246,.3)!important;transition:all .2s ease!important}
button[data-testid="baseButton-primary"]:hover{transform: translateY(-2px);box-shadow:0 6px 20px rgba(59,130,246,.4)!important}
button[data-testid="baseButton-primary"]:disabled{background:#cbd5e1!important;box-shadow:none!important;transform:none}

/* File uploader */
[data-testid="stFileUploadDropzone"]{background:#f8fafc!important;border:2px dashed #cbd5e1!important;border-radius:12px!important;transition:all .2s ease!important}
[data-testid="stFileUploadDropzone"]:hover{border-color:#3b82f6!important;background:#eff6ff!important}

/* Result card */
.rcard{background:#ffffff;border:1px solid #e2e8f0;border-radius:16px;padding:36px 32px;text-align:center;margin-top:20px;box-shadow: 0 4px 6px -1px rgba(0,0,0,.05)}
.rcard h2{font-size:24px;font-weight:700;background: linear-gradient(135deg, #0ea5e9, #3b82f6); -webkit-background-clip: text; -webkit-text-fill-color: transparent; margin:10px 0 6px}
.rcard p{font-size:15px;color:#64748b;margin-bottom:22px}

/* Progress text box */
.stream-box{background:#fff;border:1px solid #e5e7eb;border-radius:12px;padding:16px 20px;
    font-family:monospace;font-size:12px;color:#374151;max-height:200px;overflow-y:auto;margin:12px 0}

/* Sidebar text */
section[data-testid="stSidebar"] p,
section[data-testid="stSidebar"] span,
section[data-testid="stSidebar"] label{color:#374151!important}

hr{border:none!important;border-top:1px solid #e5e7eb!important;margin:26px 0!important}
.stAlert{border-radius:10px!important}
</style>
""", unsafe_allow_html=True)


# ─── API Key is now provided by user via UI ───────────────────────────────────
# ─── Clean JSON helper ────────────────────────────────────────────────────────
def _clean_json(raw: str) -> str:
    text = re.sub(r"```(?:json)?", "", raw).strip()
    s, e = text.find("{"), text.rfind("}")
    if s != -1 and e > s:
        text = text[s: e + 1]
    return re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]", "", text)


# ─── Sidebar ──────────────────────────────────────────────────────────────────
with st.sidebar:
    if os.path.exists("logo.png"):
        st.image("logo.png", use_container_width=True)
    elif os.path.exists("logo.jpg"):
        st.image("logo.jpg", use_container_width=True)
    else:
        st.markdown("## 📋 Contract Compare")
        
    st.caption("Hotel Contract Analysis · v3.1")
    st.divider()

    st.markdown("**⚙️ AI Engine**")
    api_key = st.text_input("Gemini API Key", type="password", placeholder="AIza…")
    
    if api_key:
        ok, msg = validate_api_key(api_key)
        if ok:
            st.success(msg)
        else:
            st.error(msg)
            st.info("🔑 Go to [aistudio.google.com](https://aistudio.google.com) → Get API key")
    else:
        st.warning("API Key required to run.")

    st.divider()
    st.markdown("**Scan Rules**")
    for r in [
        "100% Full Scan — No Sampling",
        "English Output Only",
        "Auto % & THB Diff (colour-coded)",
        "Exception Clause Detection",
        "Child Policy 11.99 Logic",
        "Combine Offer Detection",
        "New / Missing Room Detection",
    ]:
        st.caption(f"✓  {r}")


# ─── Hero ─────────────────────────────────────────────────────────────────────
# If logo exists, encode to base64 so we can put it inside the centered HTML hero
logo_html = ""
for ext in ["png", "jpg", "jpeg"]:
    if os.path.exists(f"logo.{ext}"):
        import base64
        with open(f"logo.{ext}", "rb") as f:
            b64 = base64.b64encode(f.read()).decode()
        logo_html = f'<img src="data:image/{ext};base64,{b64}" style="max-height:90px; margin-bottom:16px;">'
        break

st.markdown(f"""
<div class="hero">
  {logo_html}
  <div class="eyebrow">Hotel Intelligence</div>
  <div class="h1">Contract Compare</div>
  <div class="sub">Upload two hotel PDF contracts — AI delivers a full price comparison, condition changes, and policy analysis, exported to Excel in seconds.</div>
</div>
<div class="tags">
  <span class="tag">📋 100% Full Scan</span>
  <span class="tag">📊 Pricing Diff (% &amp; THB)</span>
  <span class="tag">👶 Child Policy Logic</span>
  <span class="tag">🔍 Exception Clause Detect</span>
  <span class="tag">🏨 New / Missing Room Alert</span>
  <span class="tag">📥 Export Excel</span>
</div>
""", unsafe_allow_html=True)

st.markdown("<hr>", unsafe_allow_html=True)

# ─── Upload ───────────────────────────────────────────────────────────────────
st.markdown("**Upload Contracts**")
c1, c2 = st.columns(2, gap="large")

with c1:
    st.markdown('<div class="card"><div class="c-eye">Contract 1</div><div class="c-ttl">Previous Year</div></div>', unsafe_allow_html=True)
    up1 = st.file_uploader("Contract 1", type=["pdf"], key="pdf1", label_visibility="collapsed")
    if up1:
        st.success(f"✓  {up1.name}")

with c2:
    st.markdown('<div class="card"><div class="c-eye">Contract 2</div><div class="c-ttl">New Year</div></div>', unsafe_allow_html=True)
    up2 = st.file_uploader("Contract 2", type=["pdf"], key="pdf2", label_visibility="collapsed")
    if up2:
        st.success(f"✓  {up2.name}")

st.markdown("<br>", unsafe_allow_html=True)

# ─── CTA ──────────────────────────────────────────────────────────────────────
ready = bool(up1 and up2 and api_key)
_, btn_col, _ = st.columns([1.5, 3, 1.5])
with btn_col:
    run = st.button("Compare Contracts  →", type="primary",
                    use_container_width=True, disabled=not ready)
    if not ready:
        hint = "Upload both contracts to continue" if not (up1 and up2) else "Add API Key in sidebar"
        st.markdown(f"<p style='text-align:center;color:#9ca3af;font-size:13px;margin-top:6px'>{hint}</p>",
                    unsafe_allow_html=True)

st.markdown("<hr>", unsafe_allow_html=True)


# ─── Processing ───────────────────────────────────────────────────────────────
if run:
    # 1. Extract PDF text (cached — instant if same file re-uploaded)
    with st.spinner("Reading PDFs…"):
        pdf1_bytes = up1.read()
        pdf2_bytes = up2.read()
        pdf1_text = extract_pdf_text(pdf1_bytes)
        pdf2_text = extract_pdf_text(pdf2_bytes)

    # 2. Stream AI response with live progress bar
    st.markdown("**🔍 AI Scanning Contracts…**")
    progress = st.progress(0, text="Starting analysis…")
    stream_placeholder = st.empty()

    chunks = []
    char_count = 0
    EXPECTED_CHARS = 6000  # rough estimate for progress bar

    for chunk in stream_contract_comparison(pdf1_text, pdf2_text, api_key):
        chunks.append(chunk)
        char_count += len(chunk)
        pct = min(int(char_count / EXPECTED_CHARS * 100), 99)
        progress.progress(pct, text=f"Analysing… ({char_count:,} characters received)")
        # show last 300 chars of streaming output so user sees AI working live
        live_text = "".join(chunks)[-300:]
        stream_placeholder.markdown(
            f'<div class="stream-box">{live_text}</div>', unsafe_allow_html=True
        )

    progress.progress(100, text="Done ✓")
    stream_placeholder.empty()
    progress.empty()

    result_raw = "".join(chunks)

    # 3. Quota guard
    if "429" in result_raw or "quota" in result_raw.lower():
        st.error("🚫  **API Quota Exceeded (429)**")
        st.warning(
            "Free-tier limit reached (20 req/day).\n\n"
            "**Fix:** Enable Billing at [aistudio.google.com](https://aistudio.google.com) "
            "→ 1,500 req/day + Gemini 1.5 Pro"
        )
        st.stop()

    # 4. Parse JSON
    cleaned = _clean_json(result_raw)
    try:
        data = json.loads(cleaned)
    except json.JSONDecodeError as ex:
        st.error(f"**JSON Parse Error:** {ex}")
        with st.expander("Raw AI response"):
            st.code(result_raw, language="text")
        st.stop()

    if "error" in data:
        err = data["error"]
        st.error("🚫 Quota Exceeded" if ("429" in err or "quota" in err.lower()) else f"AI Error: {err}")
        st.stop()

    # 5. Generate Excel
    try:
        excel_bytes = generate_comparison_excel(data)
    except Exception as ex:
        st.error(f"Excel generation failed: {ex}")
        with st.expander("JSON data"):
            st.json(data)
        st.stop()

    # 6. Success
    st.markdown("""
    <div class="rcard">
      <div style="font-size:40px">✅</div>
      <h2>Analysis Complete</h2>
      <p>Your comparison report is ready — download and open in Google Sheets.</p>
    </div>
    """, unsafe_allow_html=True)

    # Show recommendation banner
    recommendation = str(data.get("recommendation") or "").strip()
    if recommendation:
        if "✅" in recommendation:
            bg, border = "#f0fdf4", "#86efac"
        else:
            bg, border = "#fffbeb", "#fcd34d"
        st.markdown(
            f'<div style="background:{bg};border:2px solid {border};border-radius:12px;'
            f'padding:16px 24px;margin:16px 0;font-size:16px;font-weight:600;text-align:center">'
            f'{recommendation}</div>',
            unsafe_allow_html=True
        )

    st.markdown("<br>", unsafe_allow_html=True)
    _, dl, _ = st.columns([1.5, 3, 1.5])
    with dl:
        st.download_button(
            "⬇  Download Excel Report",
            data=excel_bytes,
            file_name="Contract_Comparison.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
    st.caption("In Google Sheets → File → Import → Upload .xlsx to keep all colours and layout.")

    with st.expander("Preview JSON data"):
        st.json(data)
