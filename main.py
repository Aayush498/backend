from fastapi import FastAPI, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
import google.generativeai as genai
from dotenv import load_dotenv
import os
import json

load_dotenv()

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Gemini setup
api_key = os.getenv("GEMINI_API_KEY")
genai.configure(api_key=api_key)
model = genai.GenerativeModel("gemini-2.5-flash")


@app.get("/")
def home():
    return {"message": "Backend Running"}


@app.post("/upload")
async def upload_file(file: UploadFile = File(...)):

    filename = file.filename.lower()

    if filename.endswith(".csv"):
        df = pd.read_csv(file.file)

    elif filename.endswith(".xlsx"):
        df = pd.read_excel(file.file)

    else:
        return {"error": "Only CSV/XLSX allowed"}


    # Basic stats
    total_rows = len(df)
    total_columns = len(df.columns)
    columns = list(df.columns)


    # Null analysis
    nulls = df.isnull().sum().to_dict()


    # Access Quality Analysis
    access_distribution = {}

    if "Access Quality" in df.columns:
        access_distribution = (
            df["Access Quality"]
            .value_counts()
            .to_dict()
        )


    # Sample rows
    sample_data = df.head(20).to_dict(orient="records")


    # AI Analysis Prompt
    prompt = f"""
You are a healthcare payer policy intelligence analyst.

Analyze this dataset deeply.

Columns:
{columns}

Dataset Sample:
{sample_data}

Provide:
1. Executive summary
2. Important patterns
3. Restriction trends
4. Access quality insights
5. Competitor opportunities
6. Risk areas
7. Recommendations
8. Market access strategy suggestions

Return clean JSON format.
"""


    response = model.generate_content(prompt)


    ai_text = response.text


    return {
        "total_rows": total_rows,
        "total_columns": total_columns,
        "columns": columns,
        "null_analysis": nulls,
        "access_distribution": access_distribution,
        "ai_analysis": ai_text,
    }