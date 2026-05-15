"""
app.py — Compare Hotel Contracts  (optimised build)
Streaming AI response · Cached PDF · Cached model · Clean light UI
"""
import os, re, json, toml, copy
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
.unified-card { background: var(--background-color); border: 1px solid var(--secondary-background-color); border-radius: 16px; padding: 24px; transition: all .3s ease; box-shadow: 0 4px 6px -1px rgba(0,0,0,.02); }
.unified-card:hover { border-color: rgba(59, 130, 246, 0.5); box-shadow: 0 10px 25px -5px rgba(0,0,0,.05); }
.c-eye { font-size: 11px; font-weight: 700; letter-spacing: .12em; text-transform: uppercase; color: #3b82f6; margin-bottom: 6px; }
.c-ttl { font-size: 22px; font-weight: 700; color: var(--text-color); margin-bottom: 24px; letter-spacing: -0.02em; }

/* Streamlit File Uploader Override */
div[data-testid="stFileUploader"] { width: 100% !important; }
div[data-testid="stFileUploader"] > section {
    background: var(--secondary-background-color) !important;
    border: 1px dashed rgba(156, 163, 175, 0.4) !important;
    border-radius: 12px !important;
    padding: 16px !important;
    transition: all 0.2s ease !important;
}
div[data-testid="stFileUploader"] > section:hover {
    border-color: #3b82f6 !important;
    background: rgba(59, 130, 246, 0.05) !important;
}
div[data-testid="stFileUploader"] small { display: none !important; }

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
# ─── Clean & Repair JSON helper ───────────────────────────────────────────────
def _clean_json(raw: str) -> str:
    # Remove markdown code blocks
    text = re.sub(r"```(?:json)?", "", raw).strip()
    
    # Extract only the JSON part
    s, e = text.find("{"), text.rfind("}")
    if s != -1 and e > s:
        text = text[s: e + 1]
    
    # Remove invalid control characters
    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]", "", text)
    
    return text

def _repair_json(text: str) -> str:
    """
    Attempt to fix common LLM JSON errors like missing commas, 
    trailing commas, and truncated responses.
    """
    if not text: return text
    
    # 1. Fix missing commas between properties/elements on new lines
    text = re.sub(r'("|\d|true|false|null|\]|\})\s*\n\s*"', r'\1,\n"', text)
    text = re.sub(r'("|\d|true|false|null|\]|\})\s*\n\s*(\{|\[)', r'\1,\n\2', text)
    
    # 2. Remove trailing commas (which json.loads hates)
    text = re.sub(r',\s*([\]\}])', r'\1', text)
    
    # 3. Handle truncation: Auto-close open braces/brackets
    # First, try to remove trailing half-written keys or values
    text = re.sub(r',?\s*"[^"]*"\s*:\s*[^,}\]]*$', '', text) # Truncated value
    text = re.sub(r',?\s*"[^"]*"\s*:\s*$', '', text)         # Truncated key
    text = re.sub(r',?\s*$', '', text)                      # Trailing comma/whitespace
    
    # Stack-based closing
    stack = []
    for char in text:
        if char == '{': stack.append('}')
        elif char == '[': stack.append(']')
        elif char in ('}', ']'):
            if stack and stack[-1] == char:
                stack.pop()
    
    # Add missing closers in reverse order
    if stack:
        text += "".join(reversed(stack))
        
    return text


# ─── Sidebar ──────────────────────────────────────────────────────────────────
with st.sidebar:
    if os.path.exists("logo.png"):
        st.image("logo.png", use_container_width=True)
    elif os.path.exists("logo.jpg"):
        st.image("logo.jpg", use_container_width=True)
    else:
        st.markdown("""
            <div style='font-size: 24px; font-weight: 800; letter-spacing: -0.03em; 
                        background: linear-gradient(90deg, var(--text-color), #3b82f6); 
                        -webkit-background-clip: text; -webkit-text-fill-color: transparent; 
                        margin-bottom: 2px;'>
                CONTRACT COMPARE
            </div>
        """, unsafe_allow_html=True)
        
    st.markdown("<div style='font-size: 12px; font-weight: 500; color: var(--text-color); opacity: 0.5; margin-bottom: 32px;'>Enterprise Contract Analysis</div>", unsafe_allow_html=True)

    st.markdown("<div style='font-size: 10px; font-weight: 700; letter-spacing: 0.1em; text-transform: uppercase; color: var(--text-color); opacity: 0.4; margin-bottom: 8px;'>AI Engine</div>", unsafe_allow_html=True)
    saved_key = load_key()
    
    # Sleek inline input instead of clunky expander
    api_key = st.text_input("GEMINI API KEY", type="password", value=saved_key, placeholder="Enter API Key...", label_visibility="collapsed")
    
    if api_key:
        if api_key != saved_key:
            save_key(api_key)
            ok, msg = validate_api_key(api_key)
            if ok:
                st.success("API Key Verified")
            else:
                st.error("Invalid API Key")
    else:
        st.caption("API Key required to run analysis.")

    st.markdown("<div style='font-size: 10px; font-weight: 700; letter-spacing: 0.1em; text-transform: uppercase; color: var(--text-color); opacity: 0.4; margin-bottom: 8px; margin-top: 32px;'>Recent Audits</div>", unsafe_allow_html=True)
    
    # Check history folder
    os.makedirs("history", exist_ok=True)
    history_files = sorted(os.listdir("history"), reverse=True)
    if not history_files:
        st.markdown("<div style='font-size: 13px; color: var(--text-color); opacity: 0.5; padding-left: 4px;'>No recent audits found.</div>", unsafe_allow_html=True)
    else:
        for hf in history_files[:8]:  # Show latest 8
            if hf.startswith("Comparison_"):
                # Clean up the name nicely
                display_name = hf.split('_vs_')[0].replace('Comparison_', '')
                if len(display_name) > 22:
                    display_name = display_name[:20] + "..."
            else:
                parts = hf.rsplit("_", 2)
                display_name = parts[0][:22] if len(parts) >= 3 else hf[:22]
            
            with open(os.path.join("history", hf), "rb") as f:
                st.download_button(
                    label=f"📄 \u00A0 {display_name}", # Added document icon and non-breaking space
                    data=f.read(),
                    file_name=hf,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=hf,
                    use_container_width=True,
                    help=f"Download full report for {display_name}"
                )




# ─── Hero ─────────────────────────────────────────────────────────────────────
# If logo exists, encode to base64 so we can put it inside the centered HTML hero
logo_html = ""
for ext in ["png", "jpg", "jpeg"]:
    if os.path.exists(f"logo.{ext}"):
        import base64
        with open(f"logo.{ext}", "rb") as f:
            b64 = base64.b64encode(f.read()).decode()
        logo_html = f'<div style="margin-bottom:16px;"><img src="data:image/{ext};base64,{b64}" style="max-height:90px; border-radius:12px; box-shadow: 0 4px 12px rgba(0,0,0,0.1);"></div>'
        break

st.markdown(f"""
<div class="hero">
  {logo_html}
  <div class="h1">HOTEL INTELLIGENCE<br>CONTRACT COMPARE</div>
</div>
<div style="height:1px; margin:10px 0 30px; background:linear-gradient(90deg, transparent, rgba(59,130,246,0.3), rgba(147,51,234,0.3), transparent); box-shadow: 0 4px 16px rgba(59,130,246,0.15), 0 1px 8px rgba(147,51,234,0.1);"></div>
""", unsafe_allow_html=True)

# ─── Upload ───────────────────────────────────────────────────────────────────
is_focus_mode = st.session_state.get("review_mode") or st.session_state.get("report_ready")

up1 = None
up2 = None

if not is_focus_mode:
    st.markdown("<div style='font-weight:700;font-size:12px;letter-spacing:0.1em;text-transform:uppercase;color:var(--text-color);opacity:0.5;margin-bottom:16px;'>Upload Contracts</div>", unsafe_allow_html=True)
    c1, c2 = st.columns(2, gap="large")
    with c1:
        st.markdown('<div style="font-size:18px; font-weight:800; letter-spacing:0.08em; text-transform:uppercase; background:linear-gradient(135deg, #0ea5e9, #3b82f6); -webkit-background-clip:text; -webkit-text-fill-color:transparent; margin-bottom:20px;">PREVIOUS CONTRACT</div>', unsafe_allow_html=True)
        up1 = st.file_uploader("Contract 1", type=["pdf"], key="pdf1", label_visibility="collapsed")
        if up1:
            st.success(f"Ready: {up1.name}")
    
    with c2:
        st.markdown('<div style="font-size:18px; font-weight:800; letter-spacing:0.08em; text-transform:uppercase; background:linear-gradient(135deg, #8b5cf6, #a855f7); -webkit-background-clip:text; -webkit-text-fill-color:transparent; margin-bottom:20px;">NEW CONTRACT</div>', unsafe_allow_html=True)
        up2 = st.file_uploader("Contract 2", type=["pdf"], key="pdf2", label_visibility="collapsed")
        if up2:
            st.success(f"Ready: {up2.name}")

    st.markdown("<br>", unsafe_allow_html=True)

cta_placeholder = st.empty()

# ─── CTA ──────────────────────────────────────────────────────────────────────
ready = bool(up1 and up2 and api_key)

if not is_focus_mode:
    # Reset state if new files are uploaded
    up1_name = up1.name if up1 else ""
    up2_name = up2.name if up2 else ""
    if "last_up1" not in st.session_state or st.session_state.last_up1 != up1_name or st.session_state.last_up2 != up2_name:
        st.session_state.started = False
        st.session_state.review_mode = False
        st.session_state.report_ready = False
        st.session_state.extracted_data = None
        st.session_state.last_up1 = up1_name
        st.session_state.last_up2 = up2_name

    with cta_placeholder.container():
        _, btn_col, _ = st.columns([1.5, 3, 1.5])
        with btn_col:
            if st.button("Compare Contracts  →", type="primary", use_container_width=True, disabled=not ready):
                st.session_state.started = True
            
            if not ready:
                hint = "Upload both contracts to continue" if not (up1 and up2) else "Add API Key in sidebar"
                st.markdown(f"<p style='text-align:center;color:#9ca3af;font-size:13px;margin-top:6px'>{hint}</p>",
                            unsafe_allow_html=True)

    st.markdown('<div style="height:1px; margin:10px 0 30px; background:linear-gradient(90deg, transparent, rgba(59,130,246,0.3), rgba(147,51,234,0.3), transparent); box-shadow: 0 4px 16px rgba(59,130,246,0.15), 0 1px 8px rgba(147,51,234,0.1);"></div>', unsafe_allow_html=True)
else:
    # FOCUS MODE is active: Show a Back/Reset button at the top instead of the uploaders
    _, reset_col, _ = st.columns([1.5, 3, 1.5])
    with reset_col:
        if st.button("← Upload Different Contracts", use_container_width=True):
            st.session_state.started = False
            st.session_state.review_mode = False
            st.session_state.report_ready = False
            st.session_state.extracted_data = None
            st.rerun()
    st.markdown("<br>", unsafe_allow_html=True)

# ─── Modal CSS (Fixed) ────────────────────────────────────────────────────────
st.markdown("""
<style>
.fixed-overlay {
    position: fixed; top: 0; left: 0; width: 100vw; height: 100vh;
    background: rgba(0, 0, 0, 0.55);
    backdrop-filter: blur(8px);
    -webkit-backdrop-filter: blur(8px);
    z-index: 999998;
    animation: overlayFadeIn 0.35s ease-out forwards;
}
.fixed-modal {
    position: fixed; top: 50%; left: 50%; transform: translate(-50%, -50%);
    border-radius: 24px; padding: 40px; width: 520px; max-width: 92vw;
    box-shadow: 0 32px 64px rgba(0,0,0,0.5), inset 0 1px 0 rgba(255,255,255,0.08) !important;
    text-align: center; z-index: 999999;
    background: rgba(15, 23, 42, 0.5) !important;
    border: 1px solid rgba(255,255,255,0.1) !important;
    color: #f1f5f9 !important;
    animation: modalSlideIn 0.4s cubic-bezier(0.16, 1, 0.3, 1) forwards;
}
.fixed-modal h3 { color: #f1f5f9 !important; }
.fixed-modal p  { color: #94a3b8 !important; }
.spinner-loader {
    border: 3px solid rgba(255,255,255,0.1); border-top: 3px solid #3b82f6; border-radius: 50%;
    width: 36px; height: 36px; animation: spin 0.8s linear infinite; margin: 0 auto 20px auto;
}
@keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
@keyframes overlayFadeIn {
    0%   { opacity: 0; }
    100% { opacity: 1; }
}
@keyframes modalSlideIn {
    0%   { opacity: 0; transform: translate(-50%, -44%) scale(0.95); }
    100% { opacity: 1; transform: translate(-50%, -50%) scale(1); }
}
</style>
""", unsafe_allow_html=True)


# ─── Processing ───────────────────────────────────────────────────────────────
if st.session_state.get("started"):
    placeholder = st.empty()
    
    # ─── Main Scan ──────────────────────────────────────────────────────────
    def render_modal(pct):
        placeholder.markdown(f"""
            <div class="fixed-overlay"></div>
            <div class="fixed-modal">
                <div class="spinner-loader" style="margin-bottom:20px;"></div>
                <h3 style="margin:0 0 8px; font-weight:700;">Analyzing Contracts...</h3>
                <p style="margin:0; font-size:14px; opacity:0.8;">Extracting policy rules and prices. Please wait.</p>
                <div style="margin-top:24px; background:var(--secondary-background-color); border-radius:10px; height:6px; overflow:hidden;">
                    <div style="background: linear-gradient(90deg, #3b82f6, #8b5cf6); width: {pct}%; height: 100%; transition: width 0.3s ease;"></div>
                </div>
                <div style="text-align:right; font-size:12px; margin-top:6px; font-weight:600; color:#3b82f6;">{pct}%</div>
            </div>
        """, unsafe_allow_html=True)
        
    render_modal(10)
    
    pdf1_bytes = up1.getvalue()
    pdf2_bytes = up2.getvalue()
    
    chunks = []
    char_count = 0
    EXPECTED_CHARS = 6000 

    for chunk in stream_contract_comparison(pdf1_bytes, pdf2_bytes, api_key):
        if chunk == "[RESET_STREAM]":
            chunks = []
            char_count = 0
            render_modal(10)
            continue
            
        chunks.append(chunk)
        char_count += len(chunk)
        pct = min(15 + int(char_count / EXPECTED_CHARS * 80), 98)
        render_modal(pct)

    render_modal(100)
    placeholder.empty()
    
    result_raw = "".join(chunks)
    st.session_state.started = False # Reset

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
    except json.JSONDecodeError:
        # Try to repair and parse again
        try:
            repaired = _repair_json(cleaned)
            data = json.loads(repaired)
        except json.JSONDecodeError as ex:
            st.error(f"**JSON Parse Error:** {ex}")
            st.info("The AI response was slightly malformed. You can download the raw response below and send it to the developer.")
            
            col1, col2 = st.columns(2)
            with col1:
                with st.expander("View Raw AI response"):
                    st.code(result_raw, language="text")
            with col2:
                st.download_button(
                    "Download Raw Response (Debug)",
                    data=result_raw,
                    file_name="raw_ai_response.txt",
                    mime="text/plain"
                )
            st.stop()

    if "error" in data:
        err = data["error"]
        st.error("Quota Exceeded" if ("429" in err or "quota" in err.lower()) else f"AI Error: {err}")
        st.stop()
        
    # Transition to review mode
    st.session_state.extracted_data = data
    st.session_state.started = False
    st.session_state.review_mode = True
    st.rerun()

# ─── 5. Review & Edit Mode ───────────────────────────────────────────────────
if st.session_state.get("review_mode"):
    st.markdown("""
        <div style="display:flex; align-items:center; gap:12px; margin-bottom: 20px;">
            <div style="background:linear-gradient(135deg, #3b82f6, #8b5cf6); border-radius:6px; padding:6px 12px; color:white; font-weight:700; font-size:12px; letter-spacing:1px; box-shadow:0 4px 6px -1px rgba(59, 130, 246, 0.3);">DATA VERIFICATION</div>
            <div style="font-size:24px; font-weight:700; color:var(--text-color); letter-spacing:-0.03em;">Review & Edit Prices</div>
        </div>
        <div style="background:var(--secondary-background-color); border-left: 4px solid #3b82f6; border-radius:4px 8px 8px 4px; padding:16px 20px; margin-bottom: 32px;">
            <p style="margin:0; font-size:14px; color:var(--text-color); opacity:0.9; line-height:1.6;">
                AI extraction is complete. Please verify the extracted prices below. You can <b>click any cell to edit</b> the value before finalizing the Excel report.
            </p>
        </div>
    """, unsafe_allow_html=True)
    
    edited_data = copy.deepcopy(st.session_state.extracted_data)
    
    # Display editable tables for seasons
    for i, season in enumerate(edited_data.get("seasons", [])):
        s_name = season.get("season_name") or f"Season {i+1}"
        p1 = season.get("period_1", "")
        p2 = season.get("period_2", "")
        
        p1_display = p1 if p1 and p1.strip() and p1 != "N/A" else "Not Specified"
        p2_display = p2 if p2 and p2.strip() and p2 != "N/A" else "Not Specified"
        
        st.markdown(f"""
            <div style="margin-top:32px; margin-bottom:16px; display:flex; align-items:baseline; flex-wrap:wrap; gap:12px; border-bottom:2px solid var(--secondary-background-color); padding-bottom:8px;">
                <div style="font-size:17px; font-weight:700; color:var(--text-color);">{s_name}</div>
                <div style="font-size:13px; color:var(--text-color);">
                    <span style="opacity:0.6;">Prev:</span> <span style="font-weight:600; opacity:0.9;">{p1_display}</span> 
                    <span style="margin:0 8px;opacity:0.2;">|</span> 
                    <span style="opacity:0.6;">New:</span> <span style="font-weight:600; color:#3b82f6;">{p2_display}</span>
                </div>
            </div>
        """, unsafe_allow_html=True)
        
        rooms = season.get("rooms", [])
        if rooms:
            edited_rooms = st.data_editor(
                rooms,
                column_config={
                    "room_name": st.column_config.TextColumn("Room Name", width="large"),
                    "price_1": st.column_config.TextColumn("Contract 1 Price"),
                    "price_2": st.column_config.TextColumn("Contract 2 Price"),
                },
                hide_index=True,
                key=f"editor_season_{i}",
                use_container_width=True
            )
            edited_data["seasons"][i]["rooms"] = edited_rooms
            
    st.markdown("<br>", unsafe_allow_html=True)
    _, btn1, btn2, _ = st.columns([1, 1.5, 1.5, 1])
    with btn1:
        if st.button("Cancel & Start Over", use_container_width=True):
            st.session_state.started = False
            st.session_state.review_mode = False
            st.session_state.extracted_data = None
            st.rerun()
    with btn2:
        if st.button("Confirm & Generate Excel", type="primary", use_container_width=True):
            st.session_state.final_data = edited_data
            st.session_state.review_mode = False
            st.session_state.report_ready = True
            st.rerun()

# ─── 6. Generate Excel ────────────────────────────────────────────────────────
if st.session_state.get("report_ready"):
    data = st.session_state.final_data
    
    try:
        excel_bytes = generate_comparison_excel(data)
        
        # Save to history folder
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        # Ensure hotel_name is a string and not empty
        hotel_name_raw = data.get("hotel_name")
        if not hotel_name_raw or not str(hotel_name_raw).strip():
            hotel_name = "Unknown_Hotel"
        else:
            hotel_name = str(hotel_name_raw).strip()
        
        if hotel_name.upper() == "HOTEL NAME":
            hotel_name = "Unknown_Hotel"
        
        hotel_name_safe = re.sub(r'[\\/*?:"<>|]', "", hotel_name)
        history_file_name = f"{hotel_name_safe}_{timestamp}.xlsx"
        
        try:
            os.makedirs("history", exist_ok=True)
            with open(os.path.join("history", history_file_name), "wb") as f:
                f.write(excel_bytes)
        except Exception as e:
            st.warning(f"Could not save to history: {e}")
            
    except Exception as ex:
        st.error(f"Excel generation failed: {ex}")
        with st.expander("Debug: JSON data"):
            st.json(data)
        st.stop()

    st.markdown("""
        <div style="background: linear-gradient(135deg, rgba(16,185,129,0.1), rgba(59,130,246,0.1));
                    border: 1px solid rgba(16,185,129,0.3); border-radius: 12px;
                    padding: 16px 24px; margin: 24px 0; display:flex; align-items:center; gap:12px;">
            <div style="width:8px;height:8px;border-radius:50%;background:#10b981;box-shadow:0 0 8px #10b981;"></div>
            <div style="font-size:14px;font-weight:600;color:var(--text-color);">Analysis complete — your comparison report is ready to download.</div>
        </div>
    """, unsafe_allow_html=True)
    
    # Show recommendation banner
    recommendation = str(data.get("recommendation") or "").strip()
    if recommendation:
        st.markdown(
            f'<div style="background:var(--secondary-background-color);border-radius:10px;'
            f'padding:14px 20px;margin:8px 0 16px;font-size:15px;font-weight:600;color:var(--text-color)">'
            f'{recommendation}</div>',
            unsafe_allow_html=True
        )

    col1, col2 = st.columns(2)
    with col1:
        st.download_button(
            "Download Excel Report",
            data=excel_bytes,
            file_name="Contract_Comparison.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
        st.caption("In Google Sheets → File → Import → Upload .xlsx")
        
    with col2:
        if st.button("Compare Another", use_container_width=True):
            st.session_state.started = False
            st.session_state.review_mode = False
            st.session_state.report_ready = False
            st.session_state.extracted_data = None
            st.rerun()


