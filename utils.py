"""
utils.py — PDF extraction + Gemini AI contract comparison
Uses the NEW google.genai SDK (google-genai package)
"""
import io
import pdfplumber
import streamlit as st
from google import genai
from google.genai import types

# ─── Fallback model list (tried in order) ───────────────────────────────────────
_FALLBACK_MODELS = [
    "gemini-2.0-flash",
    "gemini-2.0-flash-001",
    "gemini-2.0-flash-lite",
    "gemini-1.5-flash",
    "gemini-1.5-flash-latest",
    "gemini-1.5-flash-001",
    "gemini-pro",
]


# ─── Cached client (one init per session) ────────────────────────────────────
@st.cache_resource(show_spinner=False)
def _get_client(api_key: str) -> genai.Client:
    return genai.Client(api_key=api_key)


# ─── Auto-detect working model ─────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def detect_available_model(api_key: str) -> tuple[str, list[str]]:
    """
    Try ListModels first, then fall back to pinging each model.
    Returns (best_model_name, all_available_names).
    """
    client = _get_client(api_key)
    available: list[str] = []
    try:
        for m in client.models.list():
            name = getattr(m, "name", "") or ""
            short = name.replace("models/", "")
            # Skip specialized models that don't support text/document processing
            if any(skip in short.lower() for skip in ["tts", "audio", "vision", "embedding", "tuning"]):
                continue
            if any(k in short for k in ["flash", "pro"]):
                available.append(short)
    except Exception:
        pass

    # Prefer 2.0-flash, then anything with 'flash', then 'pro'
    for preferred in ["gemini-2.0-flash", "gemini-2.0-flash-lite",
                      "gemini-1.5-flash", "gemini-1.5-flash-latest"]:
        if preferred in available:
            return preferred, available

    for m in available:
        if "flash" in m:
            return m, available
    for m in available:
        if "pro" in m:
            return m, available

    # If listing failed, try pinging each fallback
    for model_name in _FALLBACK_MODELS:
        try:
            resp = client.models.generate_content(
                model=model_name,
                contents="ping",
                config=types.GenerateContentConfig(max_output_tokens=5),
            )
            if resp:
                return model_name, [model_name]
        except Exception:
            continue

    return "", available


# ─── Validate API key ─────────────────────────────────────────────────────────────
def validate_api_key(api_key: str) -> tuple[bool, str]:
    """Returns (ok, message)."""
    if not api_key:
        return False, "No API key provided."
    model, available = detect_available_model(api_key)
    if model:
        return True, f"Connected ✓  Model: {model}"
    if available:
        return True, f"Connected (models found: {len(available)})"
    return False, "API key invalid or Gemini API not enabled for this project."



# ─── Cached PDF extraction ────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def extract_pdf_text(file_bytes: bytes) -> str:
    """Extract all text from PDF bytes. Cached — re-uploads are instant."""
    pages = []
    try:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    pages.append(t)
        return "\n".join(pages)
    except Exception as e:
        return f"[PDF READ ERROR: {e}]"


# ─── Build prompt ─────────────────────────────────────────────────────────────
def _build_prompt(pdf1: str, pdf2: str) -> str:
    return f"""
You are an expert hotel contract auditor and data analyst.
Your task: compare two hotel contract PDFs strictly — Contract 1 (previous year) vs Contract 2 (new year).
Perform a 100% FULL SCAN. DO NOT skip or sample any data.
Every room type, every season/period, every fee, every policy, and every condition MUST be extracted and compared line-by-line.
Output ONLY a single valid JSON object. No markdown fences, no explanations, no extra text.

═══════════════════════════════════════════════
OPERATIONAL DIRECTIVES (STRICT — NO EXCEPTIONS)
═══════════════════════════════════════════════
1. 100% Full Scan — extract EVERY period, EVERY room, EVERY season. Never skip or summarise.
2. English only in all output fields.
3. If hotel names differ between contracts → populate the "warning" field with a bold warning message.
4. Prices MUST appear on a SEPARATE row from the period/season header row (never on the same row).
5. COMPULSORY charges (e.g. Gala Dinner, New Year's Eve Dinner) go in section2 "conditions" array ONLY. Never in section3.
6. E.B (Early Bird) / Long Stay / Benefits / Promotions MUST go in section3_conditions ONLY. Never in section1.
7. diff_amount  = price_2 − price_1  (positive = Contract 2 more expensive, prefix "+"; negative prefix "-")
8. diff_percent = (price_2 − price_1) / price_1 × 100, formatted as "+4.98%" or "-3.21%"
9. Child Policy 11.99 Rule: if contract says "under 12 yrs" → write "11.99"; if "5–11 yrs" → write "5–11.99 yrs". Never use "12" as upper bound unless the contract explicitly says "12.99".
10. Room in Contract 2 only → append "(New Room in Contract YY/YY)" to room_name.
    Room missing in Contract 2 → append "(Not have in Contract YY/YY)" to room_name.
11. Year format: "2025–2026" → "25/26". Use short format in all year fields.
12. Period date format: "1 NOV 24 - 24 DEC 24 (Season Name)".
    Multiple sub-periods: "1 Nov-22 Dec 25 / 1 Apr-14 Jul 26 / 1 Sep-31 Oct 26".
13. For each Early Bird / Long Stay offer, explicitly state whether it is combinable with other offers or not.
    If combinable → list what it can be combined with. If not combinable → state "Not combinable with other offers".
    If Black Out dates exist → state them, e.g. "*Black Out: 23 DEC 26 - 8 JAN 27".
14. Child Policy pattern:
      Child 0-3.99 yrs Sharing Bed + ABF = XX THB (or FOC)
      Child 4-11.99 yrs Extra Person + ABF = XX THB (or FOC)
      Adult Extra Person + ABF = XX THB
15. If Column F conditions overflow to other period rows, that is acceptable.
16. Section 1 covers Book-by / Stay Period promotions ONLY. Skip section1 entirely if neither contract has such promotions.
17. At the end, add "recommendation" field:
      "✅ Recommend to renew the contract" if Contract 2 is overall better.
      "🔁 Recommend key points to consider." if Contract 2 has unfavourable changes.

═══════════════════════════════════════════════
CONTRACT 1 (Previous Year):
{pdf1}

CONTRACT 2 (New Year):
{pdf2}

═══════════════════════════════════════════════
RETURN THIS EXACT JSON STRUCTURE (no extra keys, no markdown):
{{
  "hotel_name": "",
  "year_1": "",
  "year_2": "",
  "warning": "",
  "recommendation": "",
  "section1_promotions": [
    {{
      "condition": "",
      "promo1": "",
      "promo2": "",
      "diff": ""
    }}
  ],
  "section2_periods": [
    {{
      "season_name": "",
      "period_1": "",
      "period_2": "",
      "conditions": [],
      "change_note": "",
      "rooms": [
        {{
          "room_name": "",
          "price_1": 0,
          "price_2": 0,
          "diff_percent": "",
          "diff_amount": 0
        }}
      ]
    }}
  ],
  "section3_conditions": [
    {{
      "condition_type": "",
      "contract_1": "",
      "contract_2": "",
      "diff": ""
    }}
  ]
}}
"""


# ─── Streaming comparison ─────────────────────────────────────────────────────
def stream_contract_comparison(pdf1: str, pdf2: str, api_key: str):
    """
    Generator yielding text chunks from Gemini as they stream in.
    Auto-detects the best available model for this API key.
    """
    client = _get_client(api_key)
    prompt = _build_prompt(pdf1, pdf2)

    config = types.GenerateContentConfig(
        temperature=0,
        max_output_tokens=32768,
    )

    # Auto-detect best model
    best_model, all_models = detect_available_model(api_key)
    if not best_model:
        # Last resort: try every fallback
        models_to_try = _FALLBACK_MODELS
    else:
        # Try best first, then the rest of the detected list as fallback
        others = [m for m in all_models if m != best_model]
        models_to_try = [best_model] + others + _FALLBACK_MODELS

    # Deduplicate while preserving order
    seen = set()
    unique_models = []
    for m in models_to_try:
        if m not in seen:
            seen.add(m)
            unique_models.append(m)

    all_errors: list[str] = []
    for model_name in unique_models:
        try:
            for chunk in client.models.generate_content_stream(
                model=model_name,
                contents=prompt,
                config=config,
            ):
                if chunk.text:
                    yield chunk.text
            return  # success — stop trying other models
        except Exception as e:
            err = str(e)
            all_errors.append(f"{model_name}: {err}")
            # Continue to next model on ANY error
            continue

    # If we get here, ALL models failed
    summary = " | ".join(all_errors)
    safe = summary.replace('"', "'")
    yield f'{{"error":"All models failed. Details: {safe}"}}'


# ─── Non-streaming alias (backward compatible) ────────────────────────────────
def run_contract_comparison(pdf1: str, pdf2: str, api_key: str) -> str:
    """Collects all streaming chunks and returns the full JSON string."""
    return "".join(stream_contract_comparison(pdf1, pdf2, api_key))
