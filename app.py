"""
app.py — Compare Hotel Contracts  (optimised build)
Streaming AI response · Cached PDF · Cached model · Clean light UI
"""
import os, re, json, toml
import streamlit as st
from datetime import datetime
from utils import stream_contract_comparison, validate_api_key, detect_available_model
from excel_generator import generate_comparison_excel

# ─── Page config ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Contract Compare",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── CSS ─────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
html,body,[class*="css"]{font-family:'Inter',-apple-system,BlinkMacSystemFont,sans-serif!important}

/* Backgrounds (Support Dark Mode) */
.stApp, .block-container { background-color: transparent !important; }
.block-container { padding: 2rem 2.5rem 3rem !important; }
section[data-testid="stSidebar"] { border-right: 1px solid var(--secondary-background-color) !important; }

/* Sidebar clean buttons */
section[data-testid="stSidebar"] .stDownloadButton button {
    background-color: transparent !important;
    border: 1px solid transparent !important;
    color: var(--text-color) !important;
    border-radius: 8px !important;
    font-size: 13px !important;
    font-weight: 500 !important;
    justify-content: flex-start !important;
    padding: 6px 12px !important;
    transition: background-color 0.2s ease !important;
    box-shadow: none !important;
}
section[data-testid="stSidebar"] .stDownloadButton button:hover {
    background-color: var(--secondary-background-color) !important;
}

/* Hero */
.hero { text-align: center; padding: 44px 16px 28px; }
.eyebrow { font-size: 11px; font-weight: 700; letter-spacing: .12em; text-transform: uppercase; color: #3b82f6; margin-bottom: 8px; }

/* 20-Year Pro UX: Typography & Contrast */
.h1 {
    font-size: 52px; font-weight: 900; letter-spacing: -.03em; line-height: 1.05; margin-bottom: 16px;
    background: linear-gradient(135deg, #0284c7, #4f46e5, #9333ea);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    display: inline-block;
}
.sub { font-size: 17px; color: var(--text-color); opacity: 0.7; max-width: 540px; margin: 0 auto; line-height: 1.7; font-weight: 400; }

/* Upload card */
.card { background: var(--background-color); border: 1px solid var(--secondary-background-color); border-radius: 16px; padding: 26px 24px; transition: all .3s cubic-bezier(0.4, 0, 0.2, 1); box-shadow: 0 4px 6px -1px rgba(0,0,0,.05); }
.card:hover { border-color: #3b82f6; box-shadow: 0 20px 25px -5px rgba(59,130,246,.1), 0 10px 10px -5px rgba(59,130,246,.04); transform: translateY(-2px); }
.c-eye { font-size: 11px; font-weight: 700; letter-spacing: .12em; text-transform: uppercase; color: #3b82f6; margin-bottom: 8px; }
.c-ttl { font-size: 20px; font-weight: 700; color: var(--text-color); margin-bottom: 16px; }

/* Buttons */
button[data-testid="baseButton-primary"] {
    background: linear-gradient(135deg, #0ea5e9, #3b82f6) !important; color: #fff !important; border: none !important;
    border-radius: 12px !important; font-size: 16px !important; font-weight: 600 !important; padding: 12px 24px !important;
    box-shadow: 0 4px 14px rgba(59,130,246,.3) !important; transition: all .2s ease !important;
}
button[data-testid="baseButton-primary"]:hover { transform: translateY(-2px); box-shadow: 0 8px 25px rgba(59,130,246,.4) !important; }
button[data-testid="baseButton-primary"]:disabled { background: var(--secondary-background-color) !important; color: gray !important; box-shadow: none !important; transform: none; }

/* Result card */
.rcard { background: var(--background-color); border: 1px solid var(--secondary-background-color); border-radius: 16px; padding: 36px 32px; text-align: center; margin-top: 20px; box-shadow: 0 4px 6px -1px rgba(0,0,0,.05); }
.rcard h2 { font-size: 26px; font-weight: 800; color: var(--text-color); margin: 10px 0 6px; }
.rcard p { font-size: 15px; color: gray; margin-bottom: 22px; }

/* Hide top right runner */
[data-testid="stStatusWidget"] { display: none !important; }

</style>
""", unsafe_allow_html=True)

# ─── API Key Caching ──────────────────────────────────────────────────────────
KEY_FILE = ".gemini_key"
def load_key():
    if os.path.exists(KEY_FILE):
        with open(KEY_FILE, "r") as f: return f.read().strip()
    return ""

def save_key(k):
    with open(KEY_FILE, "w") as f: f.write(k.strip())
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
        st.markdown("<div style='font-size: 20px; font-weight: 600; letter-spacing: -0.02em; color: var(--text-color); margin-bottom: 4px;'>CONTRACT COMPARE</div>", unsafe_allow_html=True)
        
    st.markdown("<div style='font-size: 13px; font-weight: 400; color: var(--text-color); opacity: 0.6; margin-bottom: 24px;'>HOTEL CONTRACT ANALYSIS · V3.1</div>", unsafe_allow_html=True)

    st.markdown("<div style='font-size: 11px; font-weight: 600; letter-spacing: 0.08em; text-transform: uppercase; color: var(--text-color); opacity: 0.6; margin-bottom: 12px; margin-top: 16px;'>AI ENGINE</div>", unsafe_allow_html=True)
    saved_key = load_key()
    with st.expander("API Key Configuration", expanded=not bool(saved_key)):
        api_key = st.text_input("GEMINI API KEY", type="password", value=saved_key, placeholder="AIza…")
        
        if api_key:
            if api_key != saved_key:
                save_key(api_key)
            ok, msg = validate_api_key(api_key)
            if ok:
                st.success(msg)
            else:
                st.error(msg)
                st.info("Go to aistudio.google.com to get API key")
        else:
            st.warning("API Key required to run.")

    st.markdown("<div style='font-size: 11px; font-weight: 600; letter-spacing: 0.08em; text-transform: uppercase; color: var(--text-color); opacity: 0.6; margin-bottom: 12px; margin-top: 24px;'>HISTORY</div>", unsafe_allow_html=True)
    
    # Check history folder
    os.makedirs("history", exist_ok=True)
    history_files = sorted(os.listdir("history"), reverse=True)
    if not history_files:
        st.markdown("<div style='font-size: 13px; color: var(--text-color); opacity: 0.6;'>No previous comparisons.</div>", unsafe_allow_html=True)
    else:
        for hf in history_files[:7]:  # Show latest 7
            if hf.startswith("Comparison_"):
                display_name = hf.split('_vs_')[0].replace('Comparison_', '')[:15] + "..."
            else:
                parts = hf.rsplit("_", 2)
                if len(parts) >= 3:
                    display_name = parts[0][:25]
                else:
                    display_name = hf[:25]
            
            with open(os.path.join("history", hf), "rb") as f:
                st.download_button(
                    label=f"{display_name}",
                    data=f.read(),
                    file_name=hf,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=hf,
                    use_container_width=True,
                    help=f"Download {hf}"
                )




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
  <div class="h1">HOTEL INTELLIGENCE<br>CONTRACT COMPARE</div>
  <div class="sub">Upload two hotel PDF contracts — AI delivers a full price comparison, condition changes, and policy analysis, exported to Excel in seconds.</div>
</div>
<hr>
""", unsafe_allow_html=True)

# ─── Upload ───────────────────────────────────────────────────────────────────
st.markdown("**UPLOAD CONTRACTS**")
c1, c2 = st.columns(2, gap="large")

with c1:
    st.markdown('<div class="card"><div class="c-eye">Contract 1</div><div class="c-ttl">Previous Year</div></div>', unsafe_allow_html=True)
    up1 = st.file_uploader("Contract 1", type=["pdf"], key="pdf1", label_visibility="collapsed")
    if up1:
        st.success(f"Ready: {up1.name}")

with c2:
    st.markdown('<div class="card"><div class="c-eye">Contract 2</div><div class="c-ttl">New Year</div></div>', unsafe_allow_html=True)
    up2 = st.file_uploader("Contract 2", type=["pdf"], key="pdf2", label_visibility="collapsed")
    if up2:
        st.success(f"Ready: {up2.name}")

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
    loading_container = st.empty()
    
    with loading_container.container():
        st.markdown("""
        <style>
        .fixed-overlay {
            position: fixed; top: 0; left: 0; width: 100vw; height: 100vh;
            background: rgba(15, 23, 42, 0.65); backdrop-filter: blur(5px);
            z-index: 999998; animation: fadeIn 0.3s forwards;
        }
        .fixed-modal {
            position: fixed; top: 50%; left: 50%; transform: translate(-50%, -50%);
            border-radius: 20px; padding: 40px; width: 500px; max-width: 90vw;
            box-shadow: 0 25px 50px -12px rgba(0,0,0,0.5) !important; text-align: center;
            z-index: 999999; animation: popIn 0.3s cubic-bezier(0.16, 1, 0.3, 1) forwards;
            
            /* Fallback for light mode */
            background-color: #ffffff;
            border: 1px solid #e2e8f0;
        }
        .fixed-modal h3 { color: #0f172a; font-weight:800; margin-top:0; }
        .fixed-modal p { color: #64748b; font-size: 14px; margin-bottom: 50px; }
        
        /* Perfect Dark Mode support via media query */
        @media (prefers-color-scheme: dark) {
            .fixed-modal { background-color: #1e293b; border-color: #334155; }
            .fixed-modal h3 { color: #f8fafc; }
            .fixed-modal p { color: #94a3b8; }
        }
        [data-testid="stProgress"] {
            position: fixed !important;
            top: calc(50% + 50px) !important;
            left: 50% !important;
            transform: translate(-50%, -50%) !important;
            width: 400px !important;
            max-width: 80vw !important;
            z-index: 1000000 !important;
        }
        @keyframes fadeIn { 0% { opacity: 0; } 100% { opacity: 1; } }
        @keyframes popIn { 0% { opacity: 0; transform: translate(-50%, -45%) scale(0.95); } 100% { opacity: 1; transform: translate(-50%, -50%) scale(1); } }
        .spinner-loader {
            border: 4px solid var(--secondary-background-color);
            border-top: 4px solid #3b82f6;
            border-radius: 50%;
            width: 40px; height: 40px;
            animation: spin 1s linear infinite;
            margin: 0 auto 15px auto;
        }
        @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
        </style>
        <div class="fixed-overlay"></div>
        <div class="fixed-modal">
            <div class="spinner-loader"></div>
            <h3>Analyzing Contracts...</h3>
            <p>Extracting policy rules and price changes. Please wait.</p>
        </div>
        """, unsafe_allow_html=True)
        
        progress_bar = st.progress(5)
        
        pdf1_bytes = up1.read()
        pdf2_bytes = up2.read()

        chunks = []
        char_count = 0
        EXPECTED_CHARS = 6000  # rough estimate for progress bar

        for chunk in stream_contract_comparison(pdf1_bytes, pdf2_bytes, api_key):
            chunks.append(chunk)
            char_count += len(chunk)
            pct = min(10 + int(char_count / EXPECTED_CHARS * 85), 98)
            
            # Smooth loading progress
            progress_bar.progress(pct)

        progress_bar.progress(100)

    # Clear loading UI seamlessly
    loading_container.empty()

    result_raw = "".join(chunks)

    # 3. Quota guard
    if "429" in result_raw or "quota" in result_raw.lower():
        st.error("**API Quota Exceeded (429)**")
        st.warning(
            "Free-tier limit reached (20 req/day).\n\n"
            "**Fix:** Enable Billing at aistudio.google.com "
            "to increase quota."
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
        st.error("Quota Exceeded" if ("429" in err or "quota" in err.lower()) else f"AI Error: {err}")
        st.stop()

    # 5. Generate Excel and Save to History
    try:
        excel_bytes = generate_comparison_excel(data)
        
        # Save to history folder
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        hotel_name = str(data.get("hotel_name", "Unknown_Hotel")).strip()
        if not hotel_name or hotel_name == "HOTEL NAME":
            hotel_name = "Unknown_Hotel"
        
        hotel_name_safe = re.sub(r'[\\/*?:"<>|]', "", hotel_name)
        history_file_name = f"{hotel_name_safe}_{timestamp}.xlsx"
        
        os.makedirs("history", exist_ok=True)
        with open(os.path.join("history", history_file_name), "wb") as f:
            f.write(excel_bytes)
            
    except Exception as ex:
        st.error(f"Excel generation failed: {ex}")
        with st.expander("JSON data"):
            st.json(data)
        st.stop()

    # 6. Success Modal
    st.markdown("""
    <style>
    .fixed-overlay {
        position: fixed; top: 0; left: 0; width: 100vw; height: 100vh;
        background: rgba(15, 23, 42, 0.65); backdrop-filter: blur(5px);
        z-index: 999998; animation: fadeIn 0.3s forwards;
    }
    .fixed-modal {
        position: fixed; top: 50%; left: 50%; transform: translate(-50%, -50%);
        border-radius: 20px; padding: 40px; width: 500px; max-width: 90vw;
        box-shadow: 0 25px 50px -12px rgba(0,0,0,0.5) !important; text-align: center;
        z-index: 999999; animation: popIn 0.4s cubic-bezier(0.16, 1, 0.3, 1) forwards;
        
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
    }
    .fixed-modal h3 { color: #0f172a; font-weight:800; margin-top:0; font-size: 28px; }
    .fixed-modal p { color: #64748b; font-size: 15px; margin-bottom: 70px; }
    
    @media (prefers-color-scheme: dark) {
        .fixed-modal { background-color: #1e293b; border-color: #334155; }
        .fixed-modal h3 { color: #f8fafc; }
        .fixed-modal p { color: #94a3b8; }
    }
    
    /* Move main screen download button into the fixed modal! */
    .block-container [data-testid="stDownloadButton"] {
        position: fixed !important;
        top: calc(50% + 40px) !important;
        left: 50% !important;
        transform: translate(-50%, -50%) !important;
        width: 320px !important;
        max-width: 80vw !important;
        z-index: 1000000 !important;
    }
    </style>
    
    <div class="fixed-overlay"></div>
    <div class="fixed-modal">
        <h3>Analysis Complete</h3>
        <p>Your comparison report is ready to download.</p>
    </div>
    """, unsafe_allow_html=True)

    # Show recommendation banner
    recommendation = str(data.get("recommendation") or "").strip()
    if recommendation:
        # Move recommendation banner slightly down so it isn't blocked by the modal
        st.markdown("<br><br><br><br><br><br><br><br><br><br>", unsafe_allow_html=True)
        bg, border = "var(--background-color)", "var(--secondary-background-color)"
        st.markdown(
            f'<div style="background:{bg};border:2px solid {border};border-radius:12px;'
            f'padding:16px 24px;margin:16px 0;font-size:16px;font-weight:600;text-align:center;color:#0f172a">'
            f'{recommendation}</div>',
            unsafe_allow_html=True
        )

    # The download button is injected directly into the modal!
    st.download_button(
        "Download Excel Report",
        data=excel_bytes,
        file_name="Contract_Comparison.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )
    
    # Simple caption below the modal
    st.markdown("""
    <div style="position:fixed; top: calc(50% + 110px); left: 50%; transform: translateX(-50%); z-index:1000000; color:#64748b; font-size:12px; white-space:nowrap;">
        In Google Sheets → File → Import → Upload .xlsx to keep all colours
    </div>
    """, unsafe_allow_html=True)

    with st.expander("Preview JSON data"):
        st.json(data)
