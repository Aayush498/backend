"""
PA Policy Extraction Pipeline — PsO Indication
Local version (VS Code / terminal)
"""

import os
import json
import time
import re
from pathlib import Path

import pdfplumber
import pandas as pd
from dotenv import load_dotenv
from google import genai
from google.genai import types


# ============================================================
# CONFIG — EDIT THESE PATHS FOR YOUR LOCAL SYSTEM
# ============================================================
# BASE_DIR = Path(r"C:\Users\aayus\OneDrive\Desktop\pdfs")
BASE_DIR = Path(__file__).parent
FOLDER_PATH = BASE_DIR
OUTPUT_CSV = BASE_DIR / "result.csv"
SUBMISSIONS_FILE = BASE_DIR / "PA_Business_Rules.xlsx"
SUBMISSIONS_SHEET = "Submissions"

MODEL_NAME = "gemini-2.5-flash"


# ============================================================
# API SETUP — Load API key from .env file
# ============================================================
load_dotenv()
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
if not GEMINI_API_KEY:
    raise RuntimeError("GEMINI_API_KEY not found. Create a .env file with GEMINI_API_KEY=your_key")

client = genai.Client(api_key=GEMINI_API_KEY)
print(f"✅ Configured with {MODEL_NAME}")


# ============================================================
# PDF TEXT EXTRACTION
# ============================================================
def extract_pdf_text(pdf_path):
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = "\n".join(page.extract_text() or "" for page in pdf.pages)
        return text
    except Exception as e:
        print(f"❌ Error reading {pdf_path}: {e}")
        return ""


# ============================================================
# GEMINI EXTRACTION
# ============================================================
EXTRACTION_PROMPT = """You are an expert at extracting structured data from US health insurance Prior Authorization (PA) policy documents for psoriasis biologics.

You will be given the full text of ONE policy PDF and a TARGET BRAND. Extract values STRICTLY for the **Plaque Psoriasis (PsO)** indication of that brand only.

**IMPORTANT — Severity scope**: If the policy distinguishes between moderate-to-severe PsO and severe PsO, use ONLY the moderate-to-severe criteria. Ignore severe-only criteria.

## TARGET BRAND: {brand}

## OUTPUT — return ONLY valid JSON, no markdown, no commentary:

{{
  "Age": "<Age criterion. Allowed values: 'No' (no age restriction stated), '>=6', '>=18', '>=12', '>=4', '>=1', '>=19', '<18', or 'FDA labelled age' (if policy defers to FDA label without specifying numerical threshold). Use '>=18' for 'adults'. If policy lists requirements for TWO age groups, capture the YOUNGEST one.>",

  "Step Therapy Requirements Documented in Policy": "<Copy ALL step therapy language from the policy verbatim, covering BOTH indication/brand-specific steps AND universal criteria (criteria applying across all indications). Include phototherapy language if it appears within step statements. Preserve drug names and structure. If policy distinguishes moderate-to-severe vs severe PsO, capture ONLY moderate-to-severe. If no step therapy at all, output empty string ''>",

  "Number of Steps through Brands": "<Count of branded/biologic steps required. SEE COUNTING RULES BELOW. Output integer as string ('1', '2', '3', '4', '5') OR 'NA' if no branded steps required>",

  "Number of Steps through Generic": "<Count of non-biologic/generic steps required. SEE COUNTING RULES BELOW. Output integer as string ('1', '2', '3', '4') OR 'NA' if no generic steps required>",

  "Step through-Phototherapy": "<'Yes' if phototherapy (UVB/PUVA) is a MANDATORY required step AND NOT inside an OR statement. 'No' if policy does not require phototherapy as a mandatory step (including cases where phototherapy appears under OR — those are NOT counted as Yes). 'N/A' if policy lists no criteria at all>",

  "TB Test required": "<'Yes' if negative TB test (TST/IGRA) required before initiation, else 'No'>",

  "Quantity Limits": "<CRITICAL: Only capture text EXPLICITLY labeled as 'quantity limit', 'QL', 'Quantity Level Limit'. Do NOT capture 'dosage', 'dose', 'dosing limit'. If genuine QL text exists, copy verbatim. If QL mentioned but no specifics, output 'unspecified'. Else 'No'. Never use 'Yes'>",

  "Specialist Types": "<Specialist required for PsO. Allowed: 'Dermatologist', 'Dermatologist; Rheumatologist', 'Rheumatologist; Dermatologist', 'Appropriate Specialist', or empty string ''. Semicolon-separated if multiple>",

  "Initial Authorization Duration(in-months)": "<Integer as string, no 'Months' suffix. E.g., '12', '6', '3', '1', '24'. If PA for PsO exists but duration not specified, output 'Unspecified'. Never default>",

  "Reauthorization Duration(in-months)": "<Integer as string, no 'Months' suffix. E.g., '12', '6', '24', '36'. If reauth exists but duration not specified, output 'Unspecified'. If no reauth at all, empty string ''>",

  "Reauthorization Required": "<'Yes' if continuation/renewal criteria exist OR if reauth duration specified. Else 'No'>",

  "Reauthorization Requirements Documented in Policy": "<Verbatim text of reauth/continuation criteria. If reauth mentioned but criteria vague, output 'Unspecified'. If no reauth at all, empty string ''>"
}}

## STEP COUNTING RULES (CRITICAL — READ CAREFULLY)

### Connector resolution
- Statements with explicit **AND** → both required, sum the steps
- Statements with explicit **OR** → take the path with FEWER steps (least restrictive)
- **Statements with NO connector between them → DEFAULT TO OR** (least restrictive path)

### Union logic
- Universal criteria (applies to all indications) AND indication-specific criteria → both apply, joined by AND
- Sum the step counts from both layers

### Brand Selection / Cost-Based Restriction Clauses — MOST CRITICAL RULE

STEP 1: Search the policy for a "scope-of-policy" sentence near the top. Common patterns:
- "For plaque psoriasis, this policy applies to all members (new starts and continuation of therapy) requesting treatment with a targeted immune modulator"
- "For [indication], this policy applies to all members requesting treatment with..."
- "Note: this policy applies to..."

If such a sentence states the policy applies to ALL MEMBERS for the PsO indication (the target indication), the policy's restrictions are UNIVERSAL for PsO — they apply to EVERY PsO request regardless of which brand is being requested.

STEP 2: Search for cost-comparison / brand-selection clauses. Common patterns:
- "[Drug X] is more costly than [Drug Y]. Therefore, [Drug X] is medically necessary only for members who have a contraindication, intolerance or ineffective response to the available equivalent alternative targeted immune modulators per criteria below"
- "Aetna considers [list] to be medically necessary only for members who have ineffective response to equivalent alternatives"

If such a clause exists AND the scope sentence (Step 1) says the policy applies to all PsO members, then the cost-comparison clause's required trials APPLY TO ALL TARGET BRANDS FOR PsO — INCLUDING brands NOT listed in the restricted brand list.

Why: The scope sentence trumps the restricted-brand list. The restricted list defines WHICH brands the cost rationale applies to from a billing standpoint, but the trial-of-alternatives requirement applies to anyone seeking PsO treatment.

STEP 3: Count the required alternative trials from the cost-comparison clause as branded steps.

WORKED EXAMPLE (THIS EXACT POLICY):
Policy text snippet:
"Note: For plaque psoriasis, this policy applies to all members (new starts and continuation of therapy) requesting treatment with a targeted immune modulator..."
"Imuldosa, Otulfi, Pyzchiva, Selarsdi, Starjemza, Steqeyma, unbranded Stelara, Wezlana, and Yesintek are more costly to Aetna than other Ustekinumab products... Aetna considers [these] to be medically necessary only for members who have a contraindication, intolerance or ineffective response to the available equivalent alternative targeted immune modulators per criteria below"

Analysis for target = STELARA:
- Scope sentence: "For plaque psoriasis, this policy applies to all members" → PsO restrictions are universal.
- Cost-comparison clause requires trial of "equivalent alternative targeted immune modulators" before approval.
- Even though Stelara itself isn't in the restricted list, the universal scope means Stelara requests for PsO also require these alternative trials.
- **Yesintek (the preferred lower-cost equivalent) trial = 1 branded step**.
- Final # Brands count for Stelara: 1 (Yesintek) + any indication-specific brand steps.

### What counts
- Branded step: any specific brand name OR drug class the target belongs to (e.g., 'TNF antagonists' counts for infliximab)
- Generic step: methotrexate, cyclosporine, acitretin, sulfasalazine, leflunomide, topical agents
- A preferred ustekinumab/adalimumab product = branded step
- If a step says "try X" but doesn't specify brand/biologic → defaults to GENERIC step

### What does NOT count
- Phototherapy steps (excluded from both brand and generic counts — captured separately)
- Severity criteria (BSA %), TB test, age — not "steps"

### Worked example for indication-specific OR statements (Stelara policy)
After universal Yesintek step is counted, indication-specific criteria:
- Statement A: "previously received a biologic or targeted synthetic drug (e.g., Sotyktu, Otezla)"
- Statement B: "BSA ≥3% AND (inadequate response to phototherapy OR methotrexate/cyclosporine/acitretin OR clinical reason to avoid them)"

Analysis:
- Statement A and Statement B have no explicit connector → default to OR. Least restrictive path:
  - Path A: 1 branded step (prior biologic)
  - Path B: 0 branded + 1 generic (MTX/cyclosporine/acitretin, phototherapy excluded since under OR)
- Pick least restrictive = Path B (fewer steps).
- Combine: Universal (1 branded Yesintek) AND Indication (Path B = 1 generic).
- **Final: # Brands = 1, # Generic = 1, Step through-Phototherapy = No**

## VERBATIM TEXT RULES
- For Step Therapy, Quantity Limits, Reauth Requirements: copy EXACT text. Preserve drug names, numbers, bullets.
- Escape newlines as \\n in JSON.

## POLICY TEXT
{policy_text}

Return ONLY the JSON object."""


def extract_params(pdf_text, brand, retries=3):
    prompt = EXTRACTION_PROMPT.format(brand=brand, policy_text=pdf_text[:120000])

    for attempt in range(retries):
        try:
            response = client.models.generate_content(
                model=MODEL_NAME,
                contents=prompt,
                config=types.GenerateContentConfig(
                    temperature=0.0,
                    seed=42,
                    response_mime_type="application/json",
                ),
            )
            raw = response.text.strip()
            raw = re.sub(r"^```(?:json)?\s*|\s*```$", "", raw, flags=re.MULTILINE).strip()
            return json.loads(raw)
        except json.JSONDecodeError as e:
            print(f"  ⚠️ JSON parse fail (attempt {attempt+1}): {e}")
            time.sleep(2 ** attempt)
        except Exception as e:
            print(f"  ⚠️ API fail (attempt {attempt+1}): {e}")
            time.sleep(2 ** attempt)
    return None


# ============================================================
# ACCESS SCORE
# ============================================================
def compute_access_score(row):
    score = 100

    age = str(row.get("Age", "")).strip().lower()
    if age in (">=18",) or "adult" in age:
        score -= 12
    elif age in (">=19",):
        score -= 12
    elif age.startswith(">=") and age not in (">=18", ">=19"):
        score -= 2
    elif age == "fda labelled age":
        score -= 2
    elif age in ("no", ""):
        score -= 0

    bs = str(row.get("Number of Steps through Brands", "NA")).strip()
    if bs.isdigit():
        score -= min(int(bs), 5) * 7

    gs = str(row.get("Number of Steps through Generic", "NA")).strip()
    if gs.isdigit():
        score -= min(int(gs), 4) * 4

    if str(row.get("Step through-Phototherapy", "")).strip().lower() == "yes":
        score -= 3

    if str(row.get("TB Test required", "")).strip().lower() == "yes":
        score -= 2

    ql = str(row.get("Quantity Limits", "")).strip().lower()
    if ql in ("no", "", "na", "n/a", "nan"):
        pass
    elif ql == "unspecified":
        score -= 2
    else:
        score -= 4

    if str(row.get("Reauthorization Required", "")).strip().lower() == "yes":
        score -= 2

    step_text = str(row.get("Step Therapy Requirements Documented in Policy", "")).strip()
    brands_val = str(row.get("Number of Steps through Brands", "")).strip().lower()
    generic_val = str(row.get("Number of Steps through Generic", "")).strip().lower()

    is_empty_policy = (
        step_text in ("", "nan", "na")
        and brands_val in ("na", "nan", "")
        and generic_val in ("na", "nan", "")
    )
    if is_empty_policy:
        score = min(score, 85)

    return max(0, min(100, score))


# ============================================================
# LOAD SUBMISSIONS & BUILD TARGET PAIRS
# ============================================================
def load_target_pairs():
    sub_df = pd.read_excel(
        SUBMISSIONS_FILE,
        sheet_name=SUBMISSIONS_SHEET,
        usecols=["Filename", "Brand"]
    )
    return list(zip(sub_df["Filename"], sub_df["Brand"]))


# ============================================================
# MAIN EXTRACTION LOOP
# ============================================================
def main():
    target_pairs = load_target_pairs()[:1]
    print(f"▶ Will process {len(target_pairs)} pairs from Submissions tab")

    results = []
    pdf_text_cache = {}

    for i, (filename, brand) in enumerate(target_pairs, 1):
        print(f"\n[{i}/{len(target_pairs)}] ▶ {filename} | {brand}")
        pdf_path = FOLDER_PATH / filename

        if not pdf_path.exists():
            print(f"  ❌ File not found")
            continue

        if filename not in pdf_text_cache:
            pdf_text_cache[filename] = extract_pdf_text(pdf_path)
        text = pdf_text_cache[filename]

        if not text.strip():
            print(f"  ❌ Empty text extracted")
            continue

        extracted = extract_params(text, brand)
        if extracted is None:
            print(f"  ❌ Extraction failed after retries")
            continue

        row = {"Filename": filename, "Brand": brand, **extracted}
        row["Access Score"] = compute_access_score(row)
        results.append(row)
        print(f"  ✅ Done | Access Score: {row['Access Score']}")

        time.sleep(0.5)

    columns = [
        "Filename", "Brand", "Age",
        "Step Therapy Requirements Documented in Policy",
        "Number of Steps through Brands", "Number of Steps through Generic",
        "Step through-Phototherapy", "TB Test required", "Quantity Limits",
        "Specialist Types", "Initial Authorization Duration(in-months)",
        "Reauthorization Duration(in-months)", "Reauthorization Required",
        "Reauthorization Requirements Documented in Policy", "Access Score"
    ]

    result_df = pd.DataFrame(results)
    for col in columns:
        if col not in result_df.columns:
            result_df[col] = ""
    result_df = result_df[columns]
    result_df = result_df.fillna("").astype(str)

    # --- Enforce: no blank fields ---
    blank_defaults = {
        "Age": "No",
        "Step Therapy Requirements Documented in Policy": "NA",
        "Number of Steps through Brands": "NA",
        "Number of Steps through Generic": "NA",
        "Step through-Phototherapy": "N/A",
        "TB Test required": "No",
        "Quantity Limits": "No",
        "Specialist Types": "NA",
        "Initial Authorization Duration(in-months)": "Unspecified",
        "Reauthorization Duration(in-months)": "NA",
        "Reauthorization Required": "No",
        "Reauthorization Requirements Documented in Policy": "NA",
    }
    for col, default in blank_defaults.items():
        result_df[col] = result_df[col].replace("", default).replace("nan", default)

    # --- Reauth consistency rules ---
    for idx, r in result_df.iterrows():
        dur = r["Reauthorization Duration(in-months)"].strip()
        req_text = r["Reauthorization Requirements Documented in Policy"].strip()
        if (dur not in ("NA", "")) or (req_text not in ("NA", "")):
            result_df.at[idx, "Reauthorization Required"] = "Yes"

    mask = (result_df["Reauthorization Required"] == "Yes") & (result_df["Reauthorization Duration(in-months)"] == "NA")
    result_df.loc[mask, "Reauthorization Duration(in-months)"] = "Unspecified"

    mask = (result_df["Reauthorization Required"] == "Yes") & (result_df["Reauthorization Requirements Documented in Policy"] == "NA")
    result_df.loc[mask, "Reauthorization Requirements Documented in Policy"] = "Unspecified"

    # --- Sanity check ---
    blank_mask = (result_df == "").any(axis=1)
    if blank_mask.any():
        print(f"⚠️  {blank_mask.sum()} rows still have blanks:")
        print(result_df[blank_mask])
    else:
        print("✅ All 13 fields populated for every row")

    result_df.to_csv(OUTPUT_CSV, index=False)
    print(f"\n✅ Saved to {OUTPUT_CSV}")
    print(f"✅ Total processed: {len(result_df)} / {len(target_pairs)}")


if __name__ == "__main__":
    main()