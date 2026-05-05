"""
utils.py — PDF extraction + Gemini AI contract comparison
Uses the NEW google.genai SDK (google-genai package)
"""
import io
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


# ─── JSON Schema Definition ───────────────────────────────────────────────────
_policy_schema = types.Schema(
    type=types.Type.OBJECT,
    properties={
        "topic": types.Schema(type=types.Type.STRING),
        "contract_1": types.Schema(type=types.Type.ARRAY, items=types.Schema(type=types.Type.STRING)),
        "contract_2": types.Schema(type=types.Type.ARRAY, items=types.Schema(type=types.Type.STRING)),
        "diff_summary": types.Schema(type=types.Type.STRING),
    },
    required=["topic", "contract_1", "contract_2", "diff_summary"]
)

_response_schema = types.Schema(
    type=types.Type.OBJECT,
    properties={
        "step_by_step_analysis": types.Schema(type=types.Type.ARRAY, items=types.Schema(type=types.Type.STRING)),
        "hotel_name": types.Schema(type=types.Type.STRING),
        "year_1": types.Schema(type=types.Type.STRING),
        "year_2": types.Schema(type=types.Type.STRING),
        "seasons": types.Schema(
            type=types.Type.ARRAY,
            items=types.Schema(
                type=types.Type.OBJECT,
                properties={
                    "season_name": types.Schema(type=types.Type.STRING),
                    "period_1": types.Schema(type=types.Type.STRING),
                    "period_2": types.Schema(type=types.Type.STRING),
                    "conditions": types.Schema(type=types.Type.ARRAY, items=types.Schema(type=types.Type.STRING)),
                    "rooms": types.Schema(
                        type=types.Type.ARRAY,
                        items=types.Schema(
                            type=types.Type.OBJECT,
                            properties={
                                "room_name": types.Schema(type=types.Type.STRING),
                                "price_1": types.Schema(type=types.Type.STRING),
                                "price_2": types.Schema(type=types.Type.STRING),
                            },
                            required=["room_name", "price_1", "price_2"]
                        )
                    )
                },
                required=["season_name", "period_1", "period_2", "conditions", "rooms"]
            )
        ),
        "extra_bed": types.Schema(type=types.Type.ARRAY, items=_policy_schema),
        "early_bird": types.Schema(type=types.Type.ARRAY, items=_policy_schema),
        "bonus_night": types.Schema(type=types.Type.ARRAY, items=_policy_schema),
        "wellbeing": types.Schema(type=types.Type.ARRAY, items=_policy_schema),
        "cancellation": types.Schema(type=types.Type.ARRAY, items=_policy_schema),
    },
    required=[
        "step_by_step_analysis", "hotel_name", "year_1", "year_2", 
        "seasons", "extra_bed", "early_bird", "bonus_night", "wellbeing", "cancellation"
    ]
)


# ─── Build prompt ─────────────────────────────────────────────────────────────
def _build_prompt() -> str:
    return f"""
You are an expert hotel contract data extractor.
Your task is to extract data from TWO hotel contract PDFs (Contract 1: Old Year, Contract 2: New Year) and prepare it for our Python system.

Perform a 100% FULL SCAN. DO NOT skip any data, but you MUST format the text output according to the STRICT PATTERNS below. Do not output raw paragraphs.
CRITICAL JSON RULES:
1. You must output STRICTLY VALID JSON.
2. DO NOT use double quotes (") inside any of your string values. If you need to quote something, use single quotes (').
3. If you need newlines within a string, escape them as \\n. DO NOT use literal unescaped newlines.

═══════════════════════════════════════════════
FORMATTING PATTERNS (STRICT!)
═══════════════════════════════════════════════
1. **Season Conditions (for periods/seasons)**:
   DO NOT copy paragraphs. Extract and format ONLY the key points as an array of strings using the '•' bullet point:
   - • MIN. [x] NIGHTS
   - • COMPULSORY [GALA DINNER / NEW YEAR] on [Date] = [Price] THB
   - • **NOT ALLOWED CHECK OUT on [Date] - [Date]**
   - • FULL MOON Period [Date] - [Date] = [Price] THB

2. **Extra bed / Extra person / Child Policy**:
   Summarize using this exact format with the '•' bullet point:
   - • CHD ([Age]-[Age] yrs) Sharing bed + ABF = [Price or FOC]
   - • CHD ([Age]-[Age] yrs) Extra bed + ABF = [Price] THB
   - • Adult ([Age] yrs and above) Extra bed + ABF = [Price] THB

3. **Early Bird Offer & Bonus Nights & Benefits**:
   Format strictly using the '•' bullet point for titles and items:
   - • [Topic or Benefit Name]:
   - • Valid Period: [Date - Date]
   - • E.B [x] DAYS, [x]% discount.
   - • [Can combine with: Promotion Name]

4. **Cancellation Policy**:
   Format strictly using the '•' bullet point for the valid period, but NO bullet for the penalty details:
   - • Valid Period: [Season Name / Date]
   - Notice [x] days prior to arrival, penalty [x]% (or x nights)

5. **Important Notes & Remarks**:
   If there are important notes, remarks, exceptions, or conditions (e.g. "NOTE: If reservation overlaps different seasons..."), you MUST prefix that specific string with `[RED] ` so our system can color it red.
   - [RED] • NOTE: Early bird offer is NOT applicable for meal plans.

6. **Rooms & Prices**:
   Extract all room names and prices exactly. DO NOT calculate percentage or diffs! The Python system will do the math.

═══════════════════════════════════════════════
CHAIN OF THOUGHT (STRICT EXTRACTION LOGIC)
═══════════════════════════════════════════════
Mentally follow these steps:
1. Scan for Hotel Name and Contract Years.
2. Identify all Seasons/Periods. Pair them logically (e.g. 1 Nov 24 pairs with 1 Nov 25).
3. Extract rooms and prices for each season.
4. Extract Season Conditions, Extra Bed, Early Bird, Bonus Night, Wellbeing, and Cancellation, strictly converting them into the FORMATTING PATTERNS above.
5. Finalize JSON.

═══════════════════════════════════════════════
THE PDF FILES HAVE BEEN PROVIDED AS INLINE DOCUMENTS.
Contract 1 is the first PDF provided.
Contract 2 is the second PDF provided.
═══════════════════════════════════════════════
RETURN THIS EXACT JSON STRUCTURE:
{{
  "step_by_step_analysis": [
    "Thinking step 1: I am looking at Room A...",
    "Thinking step 2: Comparing conditions..."
  ],
  "hotel_name": "Hotel Name",
  "year_1": "24/25",
  "year_2": "25/26",
  "seasons": [
    {{
      "season_name": "PEAK SEASON (or empty)",
      "period_1": "1 NOV 24 - 23 DEC 24",
      "period_2": "1 NOV 25 - 23 DEC 25",
      "conditions": [
        "MIN. 3 NIGHTS",
        "COMPULSORY GALA DINNER on 31 DEC = 1000 THB"
      ],
      "rooms": [
        {{
          "room_name": "Deluxe Room",
          "price_1": 15000,
          "price_2": 16000
        }}
      ]
    }}
  ],
  "extra_bed": [
    {{
      "topic": "Child / Extra bed policy",
      "contract_1": ["CHD (4-11.99 yrs) Sharing bed = FOC", "Adult Extra bed = 1500 THB"],
      "contract_2": ["CHD (4-11.99 yrs) Sharing bed = FOC", "Adult Extra bed = 1500 THB"]
    }}
  ],
  "early_bird": [
    {{
      "topic": "Early Bird details",
      "contract_1": ["• Valid Period: 1 Nov 24 - 31 Oct 25", "• E.B 60 DAYS, 10% discount", "• Blackout dates: None", "[RED] • NOTE: Not applicable for meal plans"],
      "contract_2": ["• Valid Period: 1 Nov 25 - 31 Oct 26", "• E.B 60 DAYS, 10% discount", "• Blackout dates: None", "[RED] • NOTE: Not applicable for meal plans"],
      "diff_summary": "SAME"
    }}
  ],
  "bonus_night": [
    {{
      "topic": "Bonus night details",
      "contract_1": ["• Valid Period: 1 May 24 - 31 Oct 24", "• STAY 5 PAY 4"],
      "contract_2": ["• Valid Period: 1 May 25 - 31 Oct 25", "• STAY 5 PAY 4"],
      "diff_summary": "SAME"
    }}
  ],
  "wellbeing": [
    {{
      "topic": "Long stay / Wellbeing",
      "contract_1": ["• Honeymoon Benefits (Minimum 3 nights stay required):", "• Welcome Drink and Cold Towel upon arrival", "• Complimentary one (1) slice of cake in room"],
      "contract_2": ["• Honeymoon Benefits (Minimum 3 nights stay required):", "• Welcome Drink and Cold Towel upon arrival", "• Complimentary one (1) slice of cake in room"],
      "diff_summary": "SAME"
    }}
  ],
  "cancellation": [
    {{
      "topic": "Cancellation Policy",
      "contract_1": ["• Valid Period: Peak Season", "Notice 45 days prior to arrival, penalty 100% of total booking revenue"],
      "contract_2": ["• Valid Period: Peak Season", "Notice 45 days prior to arrival, penalty 100% of total booking revenue", "[RED] • NOTE: If reservation overlaps different seasons, the higher season cancellation terms apply"],
      "diff_summary": "25-26 added Note for overlapping seasons"
    }}
  ]
}}
"""


# ─── Streaming comparison ─────────────────────────────────────────────────────
def stream_contract_comparison(pdf1_bytes: bytes, pdf2_bytes: bytes, api_key: str):
    """
    Generator yielding text chunks from Gemini as they stream in.
    Auto-detects the best available model for this API key.
    """
    client = _get_client(api_key)
    prompt = _build_prompt()

    config = types.GenerateContentConfig(
        temperature=0,
        top_p=0.0,
        top_k=1,
        max_output_tokens=32768,
        response_mime_type="application/json",
        response_schema=_response_schema,
    )

    # Convert PDFs to native Gemini Parts
    pdf1_part = types.Part.from_bytes(data=pdf1_bytes, mime_type="application/pdf")
    pdf2_part = types.Part.from_bytes(data=pdf2_bytes, mime_type="application/pdf")
    
    # Construct multi-modal payload
    contents = [
        "Please analyze the following two hotel contracts.",
        "\n\n--- CONTRACT 1 (Previous Year) ---\n",
        pdf1_part,
        "\n\n--- CONTRACT 2 (New Year) ---\n",
        pdf2_part,
        "\n\n--- INSTRUCTIONS ---\n",
        prompt
    ]

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
                contents=contents,
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
def run_contract_comparison(pdf1_bytes: bytes, pdf2_bytes: bytes, api_key: str) -> str:
    """Collects all streaming chunks and returns the full JSON string."""
    return "".join(stream_contract_comparison(pdf1_bytes, pdf2_bytes, api_key))
