import os
import re
import PyPDF2
import docx2txt
from flask import Flask, request, render_template, redirect, send_file
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
import pandas as pd

# ---------- FLASK APP ----------
app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
RESUME_FOLDER = os.path.join(BASE_DIR, "resumes")
JD_FOLDER = os.path.join(BASE_DIR, "job_description")

# ---------- HELPER FUNCTIONS ----------
def clean_text(text):
    text = text.lower()
    text = re.sub(r'[^a-z0-9\s]', ' ', text)
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

def extract_pdf_text(file_path):
    text = ""
    with open(file_path, "rb") as file:
        reader = PyPDF2.PdfReader(file)
        for page in reader.pages:
            text += page.extract_text()
    return clean_text(text)

def extract_docx_text(file_path):
    return clean_text(docx2txt.process(file_path))

def extract_job_description(file_path):
    with open(file_path, "r", encoding="utf-8") as f:
        return clean_text(f.read())

def calculate_match_score(jd_text, resume_text):
    vectorizer = TfidfVectorizer()
    vectors = vectorizer.fit_transform([jd_text, resume_text])
    score = cosine_similarity(vectors[0], vectors[1])[0][0]
    return round(score * 100, 2)

# ---------- ROUTES ----------
@app.route("/", methods=["GET", "POST"])
def index():
    results = []
    if request.method == "POST":
        # Save uploaded JD
        jd_file = request.files["jd"]
        jd_path = os.path.join(JD_FOLDER, "jd.txt")
        jd_file.save(jd_path)

        # Save uploaded resumes
        uploaded_files = request.files.getlist("resumes")
        for file in uploaded_files:
            save_path = os.path.join(RESUME_FOLDER, file.filename)
            file.save(save_path)

        # Extract JD text
        jd_text = extract_job_description(jd_path)

        # Calculate scores for all resumes
        for resume_file in os.listdir(RESUME_FOLDER):
            path = os.path.join(RESUME_FOLDER, resume_file)
            if resume_file.endswith(".pdf"):
                text = extract_pdf_text(path)
            elif resume_file.endswith(".docx"):
                text = extract_docx_text(path)
            else:
                continue
            score = calculate_match_score(jd_text, text)
            results.append({"name": resume_file, "score": score})

        # Sort by score
        results = sorted(results, key=lambda x: x["score"], reverse=True)

        # Save to Excel
        df = pd.DataFrame(results)
        df.to_excel(os.path.join(BASE_DIR, "results.xlsx"), index=False)

    return render_template("index.html", results=results)

# ---------- DOWNLOAD ROUTE ----------
@app.route("/download")
def download():
    file_path = os.path.join(BASE_DIR, "results.xlsx")
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    else:
        return "No results file found. Run AI Screening first."

# ---------- RUN APP ----------
if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
