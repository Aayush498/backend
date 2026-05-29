# PA Policy Extraction Pipeline — Local Version

End-to-end pipeline that extracts structured Prior Authorization (PA) policy data from psoriasis biologic PDFs and computes an **Access Score** for each filename–brand pair.

This README covers the **local (`local_version.py`)** flow — runs on your machine via VS Code or terminal using a `.env` file for the API key.

For Google Colab, see `README_collab.md`.

---

## Table of Contents

1. [What this pipeline does](#1-what-this-pipeline-does)
2. [Architecture & flow](#2-architecture--flow)
3. [Prerequisites](#3-prerequisites)
4. [Folder structure](#4-folder-structure)
5. [Setup — step by step](#5-setup--step-by-step)
6. [How to run](#6-how-to-run)
7. [Output schema (13 fields + Access Score)](#7-output-schema-13-fields--access-score)
8. [Access Score formula](#8-access-score-formula)
9. [Blank-field guarantee](#9-blank-field-guarantee)
10. [Troubleshooting](#10-troubleshooting)
11. [Cost & runtime](#11-cost--runtime)
12. [FAQ](#12-faq)

---

## 1. What this pipeline does

For each `(Filename, Brand)` pair listed in `PA_Business_Rules.xlsx` (Submissions tab):

1. Opens the corresponding PDF from your local folder
2. Extracts all text using `pdfplumber`
3. Sends the text + target brand to **Gemini 2.5 Flash** with a strict extraction prompt
4. Receives a JSON object containing 13 structured fields about the policy
5. Computes a numeric **Access Score** (0–100) using a deterministic formula
6. Enforces blank-field defaults and reauthorization consistency rules
7. Writes the final result to `result.csv`

The output is a CSV with one row per `(Filename, Brand)` pair and 15 columns total (13 extracted + Filename + Brand + Access Score).

---

## 2. Architecture & flow

```
PA_Business_Rules.xlsx (Submissions tab)
            │
            ▼
   load_target_pairs()  →  [(file1, brand1), (file2, brand2), ...]
            │
            ▼
    For each pair:
       ├── extract_pdf_text()       → raw text from PDF
       ├── extract_params()         → Gemini API call → 13-field JSON
       └── compute_access_score()   → 0–100 numeric score
            │
            ▼
   blank_defaults  →  fill any empty fields with semantic defaults
            │
            ▼
   Reauth consistency rules
            │
            ▼
        result.csv
```

PDF text is cached in-memory across the loop so the same PDF is read only once even if multiple brands map to it.

---

## 3. Prerequisites

- **Python**: 3.9 or higher
- **OS**: Windows / macOS / Linux
- **Gemini API key**: get one from https://aistudio.google.com/apikey
- **Local PDFs**: all PDFs referenced in `PA_Business_Rules.xlsx` must exist locally
- **Excel file**: `PA_Business_Rules.xlsx` with a sheet named `Submissions` containing columns `Filename` and `Brand`

---

## 4. Folder structure

Recommended layout:

```
your_project/
├── local_version.py
├── .env                          ← your API key (DO NOT commit)
├── README_local.md
└── pdfs/                         ← BASE_DIR points here
    ├── PA_Business_Rules.xlsx    ← Submissions sheet inside
    ├── result.csv                ← created after run
    ├── 313179-3560271.pdf
    ├── 296961-4569911.pdf
    └── ... (all other PDFs)
```

`BASE_DIR` in the script must point to the folder that contains both the PDFs and `PA_Business_Rules.xlsx`. The output `result.csv` is written into the same folder.

---

## 5. Setup — step by step

### 5.1 Clone or download the script

Save `local_version.py` somewhere on your machine.

### 5.2 Install Python dependencies

Open a terminal in the folder containing `local_version.py` and run:

```bash
pip install pdfplumber google-genai pandas openpyxl python-dotenv
```

Package roles:
- `pdfplumber` — PDF text extraction
- `google-genai` — Gemini API client
- `pandas` — DataFrame + Excel/CSV I/O
- `openpyxl` — Excel `.xlsx` engine for pandas
- `python-dotenv` — load API key from `.env`

### 5.3 Create the `.env` file

In the same folder as `local_version.py`, create a file named `.env` (no extension before the dot) with this single line:

```
GEMINI_API_KEY=your_actual_key_here
```

Do not put quotes around the key. Do not commit `.env` to git.

### 5.4 Set up the PDF folder

Create a folder (e.g. `C:\Users\<you>\OneDrive\Desktop\pdfs` or `~/pdfs`) and place inside it:
- All PDF files referenced in the Submissions sheet
- `PA_Business_Rules.xlsx` with a `Submissions` sheet containing `Filename` and `Brand` columns

### 5.5 Update `BASE_DIR` in the script

Open `local_version.py` and edit this line near the top:

```python
BASE_DIR = Path(r"C:\Users\aayus\OneDrive\Desktop\pdfs")
```

Change to your actual local path. On macOS / Linux:

```python
BASE_DIR = Path("/Users/yourname/pdfs")
```

The `r"..."` prefix is for Windows paths to handle backslashes; on Unix you can drop it.

---

## 6. How to run

From your terminal in the folder containing `local_version.py`:

```bash
python local_version.py
```

You should see:

```
✅ Configured with gemini-2.5-flash
▶ Will process 79 pairs from Submissions tab

[1/79] ▶ 313179-3560271.pdf | STELARA
  ✅ Done | Access Score: 57

[2/79] ▶ 313179-3560271.pdf | TREMFYA
  ✅ Done | Access Score: 61
...

✅ All 13 fields populated for every row

✅ Saved to C:\Users\...\pdfs\result.csv
✅ Total processed: 79 / 79
```

If the script can't find a PDF, it prints `❌ File not found` and skips that pair. If the Gemini API fails for a pair, it retries up to 3 times with exponential backoff, then skips with `❌ Extraction failed after retries`.

---

## 7. Output schema (13 fields + Access Score)

| # | Column | Type | Allowed values / format |
|---|---|---|---|
| 1 | `Filename` | string | PDF filename |
| 2 | `Brand` | string | Target brand name |
| 3 | `Age` | string | `No`, `>=6`, `>=18`, `>=12`, `>=4`, `>=1`, `>=19`, `<18`, `FDA labelled age` |
| 4 | `Step Therapy Requirements Documented in Policy` | string | Verbatim text or `NA` |
| 5 | `Number of Steps through Brands` | string | `1`–`5` or `NA` |
| 6 | `Number of Steps through Generic` | string | `1`–`4` or `NA` |
| 7 | `Step through-Phototherapy` | string | `Yes`, `No`, `N/A` |
| 8 | `TB Test required` | string | `Yes`, `No` |
| 9 | `Quantity Limits` | string | Verbatim QL text, `unspecified`, or `No` |
| 10 | `Specialist Types` | string | `Dermatologist`, `Dermatologist; Rheumatologist`, `Appropriate Specialist`, etc. |
| 11 | `Initial Authorization Duration(in-months)` | string | Integer (e.g. `12`) or `Unspecified` |
| 12 | `Reauthorization Duration(in-months)` | string | Integer, `Unspecified`, or `NA` |
| 13 | `Reauthorization Required` | string | `Yes`, `No` |
| 14 | `Reauthorization Requirements Documented in Policy` | string | Verbatim text, `Unspecified`, or `NA` |
| 15 | `Access Score` | integer | 0–100 |

Numeric fields are stored as **strings** to preserve `NA` / `Unspecified` semantics. Convert downstream as needed.

---

## 8. Access Score formula

Start at 100, subtract penalties:

| Criterion | Penalty |
|---|---|
| `Age` = `>=18` / `>=19` / contains "adult" | −12 |
| `Age` = `>=6`, `>=4`, `>=12`, etc. (pediatric-inclusive) | −2 |
| `Age` = `FDA labelled age` | −2 |
| `Number of Steps through Brands` (capped at 5) | −7 each |
| `Number of Steps through Generic` (capped at 4) | −4 each |
| `Step through-Phototherapy` = `Yes` | −3 |
| `TB Test required` = `Yes` | −2 |
| `Quantity Limits` = explicit text | −4 |
| `Quantity Limits` = `unspecified` | −2 |
| `Reauthorization Required` = `Yes` | −2 |

**Empty-policy cap**: if step therapy is blank AND both brand/generic step counts are `NA`, the score is capped at 85 (to flag incomplete extractions, not reward them as perfect-access policies).

**Bounds**: final score is clamped to `[0, 100]`.

---

## 9. Blank-field guarantee

Submission rule: *all 13 fields must be populated for every row*.

After extraction, the pipeline applies `blank_defaults` to fill any empty cell with a semantically correct placeholder:

| Field | Default when blank |
|---|---|
| Age | `No` |
| Step Therapy Requirements | `NA` |
| Number of Steps through Brands | `NA` |
| Number of Steps through Generic | `NA` |
| Step through-Phototherapy | `N/A` |
| TB Test required | `No` |
| Quantity Limits | `No` |
| Specialist Types | `NA` |
| Initial Authorization Duration | `Unspecified` |
| Reauthorization Duration | `NA` |
| Reauthorization Required | `No` |
| Reauthorization Requirements | `NA` |

**Reauth consistency rules** (enforced after blank-fill):
1. If `Reauthorization Duration` ≠ `NA` OR `Reauthorization Requirements` ≠ `NA` → `Reauthorization Required` is forced to `Yes`
2. If `Reauthorization Required` = `Yes` but `Duration` is `NA` → set to `Unspecified`
3. If `Reauthorization Required` = `Yes` but `Requirements` is `NA` → set to `Unspecified`

The final sanity check prints either `✅ All 13 fields populated for every row` or lists the rows that still have blanks.

---

## 10. Troubleshooting

**`GEMINI_API_KEY not found`**
The `.env` file is missing, not in the same folder as `local_version.py`, or the variable name is wrong. The file must literally contain `GEMINI_API_KEY=...` with no spaces.

**`File not found` for every PDF**
`BASE_DIR` is pointing to the wrong folder. Print it out: `print(BASE_DIR)` and verify the path exists.

**`ModuleNotFoundError: No module named 'google.genai'`**
Install with `pip install google-genai` (note the dash). Old name `google-generativeai` is different.

**`JSON parse fail` repeatedly for one file**
Gemini occasionally returns malformed JSON. The script retries 3× with backoff. If it still fails, the file is skipped — re-run later or test that specific pair in isolation.

**Garbled bullet characters in step therapy text (`â€¢`)**
This is a UTF-8 encoding issue when viewing the CSV in Excel. The data is correct; open the CSV with UTF-8 encoding or use pandas to load it instead.

**Some `Number of Steps through Brands` look wrong**
The extraction prompt is deterministic (`temperature=0.0, seed=42`) but Gemini still has interpretation variance. Cross-check with the original policy PDF for the specific row.

**Run is slow**
~1 second per pair for the PDF parse + ~3–8 seconds for the Gemini call. Total runtime for 79 pairs is roughly 7–12 minutes depending on PDF size and API latency.

---

## 11. Cost & runtime

- **Model**: `gemini-2.5-flash`
- **Input tokens per call**: ~10k–40k (depending on PDF size; capped at 120k chars)
- **Output tokens per call**: ~500–2000
- **Estimated cost**: ~$0.05–$0.15 for a full 79-pair run (current Gemini Flash pricing)
- **Runtime**: 7–12 minutes for 79 pairs

PDF text is cached so multi-brand files are only parsed once.

---

## 12. FAQ

**Can I re-run on just a subset of files?**
Yes — temporarily hardcode `target_pairs` in `main()`:

```python
target_pairs = load_target_pairs()
target_pairs = [
    ("313179-3560271.pdf", "STELARA"),
    ("296961-4569911.pdf", "TREMFYA"),
]
```

Don't forget to remove the override before a full production run.

**Does this overwrite my previous `result.csv`?**
Yes. To preserve old results, change `OUTPUT_CSV` at the top of the script:

```python
OUTPUT_CSV = BASE_DIR / "result_v2.csv"
```

**Can I change the scoring weights?**
Yes — edit `compute_access_score()`. Penalties, caps, and the empty-policy guard are all local constants.

**Why are number fields stored as strings?**
Because `NA` / `Unspecified` are valid values per the schema. Mixing integers and strings would force pandas to use `object` dtype anyway, so all fields are normalized to strings.

**Where is the prompt logic?**
The `EXTRACTION_PROMPT` constant in `local_version.py` is the full Gemini prompt — including step-counting rules, severity scope, and the Yesintek worked example. Tune that text if extraction logic needs adjustment.

**Is there any hardcoding tied to specific PDFs?**
No. All filename–brand pairs come from `PA_Business_Rules.xlsx → Submissions`. The only PDF-specific content in the codebase is the Yesintek worked example inside the prompt, which is illustrative reasoning, not hardcoded data.
