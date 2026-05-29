# PA Policy Extraction — Local Run

## 1. Folder setup

Create a folder anywhere (e.g. `C:\Users\<you>\Desktop\pdfs`) and put inside it:

```
pdfs/
├── local_version.py
├── .env
├── PA_Business_Rules.xlsx     ← must have "Submissions" sheet with Filename, Brand columns
├── 313179-3560271.pdf
├── 296961-4569911.pdf
└── ... (all other PDFs)
```

## 2. Install dependencies

Open terminal in the folder and run:

```bash
pip install pdfplumber google-genai pandas openpyxl python-dotenv
```

## 3. Add API key

Create a file named `.env` in the same folder with this line:

```
GEMINI_API_KEY=your_actual_key_here
```

Get key from: https://aistudio.google.com/apikey

## 4. Run

```bash
python local_version.py
```

Output saves to `result.csv` in the same folder.
