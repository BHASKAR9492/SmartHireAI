import os
from flask import Flask, request, render_template, redirect, send_file, url_for, flash
import pandas as pd
from werkzeug.utils import secure_filename
from PyPDF2 import PdfReader
import docx2txt

UPLOAD_FOLDER = "resumes"
JD_FILE = "job_description/jd.txt"
RESULT_FILE = "results.xlsx"
ALLOWED_EXTENSIONS = {"pdf", "docx"}

ADMIN_PASSWORD = "admin123"
RESULTS_TOKEN = "Ab12Xy9Q"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs("job_description", exist_ok=True)

app = Flask(__name__)
app.secret_key = "smarthireai-secret"
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

def extract_text(file_path):
    if file_path.endswith(".pdf"):
        reader = PdfReader(file_path)
        return " ".join([page.extract_text() for page in reader.pages if page.extract_text()])
    elif file_path.endswith(".docx"):
        return docx2txt.process(file_path)
    return ""

def extract_skills(text):
    SKILLS = {"python","java","sql","c++","excel","power bi","tableau","machine learning","deep learning","data analysis"}
    text_words = set(word.lower() for word in text.split())
    return SKILLS.intersection(text_words)

def calculate_scores(jd_text, resume_texts, resume_names):
    results = []
    jd_skills = extract_skills(jd_text)
    for name, resume_text in zip(resume_names, resume_texts):
        resume_skills = extract_skills(resume_text)
        matched_skills = jd_skills.intersection(resume_skills)
        missing_skills = jd_skills - resume_skills
        if len(jd_skills) == 0:
            score = 0
        else:
            score = (len(matched_skills) / len(jd_skills)) * 100
        results.append({
            "name": name,
            "score": round(score,2),
            "matched_skills": ", ".join(matched_skills),
            "missing_skills": ", ".join(missing_skills)
        })
    return results

def save_results(results):
    df = pd.DataFrame(results)
    df.to_excel(RESULT_FILE, index=False)

@app.route("/", methods=["GET","POST"])
def index():
    if request.method == "POST":
        files = request.files.getlist("resumes")
        if not files or files[0].filename == "":
            flash("No files selected", "danger")
            return redirect(request.url)

        resume_texts = []
        resume_names = []
        for file in files:
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
                file.save(path)
                resume_names.append(filename)
                resume_texts.append(extract_text(path))

        if os.path.exists(JD_FILE):
            with open(JD_FILE, "r", encoding="utf-8") as f:
                jd_text = f.read()
        else:
            flash("Job Description not found. Ask admin to upload.", "warning")
            return redirect(request.url)

        results = calculate_scores(jd_text, resume_texts, resume_names)
        save_results(results)
        flash("Results saved to results.xlsx", "success")
        return render_template("index.html", results=results)

    return render_template("index.html", results=None)

@app.route("/download")
def download_results():
    if os.path.exists(RESULT_FILE):
        return send_file(RESULT_FILE, as_attachment=True)
    else:
        flash("Results file not found", "danger")
        return redirect(url_for("index"))

@app.route("/admin", methods=["GET","POST"])
def admin_panel():
    if request.method == "POST":
        password = request.form.get("password")
        if password != ADMIN_PASSWORD:
            flash("Invalid password", "danger")
            return redirect(request.url)
        jd_file = request.files.get("jd_file")
        if jd_file and jd_file.filename != "":
            jd_file.save(JD_FILE)
            flash("Job Description updated", "success")
        if os.path.exists(RESULT_FILE):
            df = pd.read_excel(RESULT_FILE)
            total_candidates = len(df)
            shortlisted = df[df["score"]>=60].to_dict(orient="records")
            all_candidates = df.to_dict(orient="records")
        else:
            total_candidates = 0
            shortlisted = []
            all_candidates = []
        return render_template("admin.html", total_candidates=total_candidates,
                               shortlisted=shortlisted, all_candidates=all_candidates)
    return render_template("admin_login.html")

@app.route(f"/results/{RESULTS_TOKEN}")
def results_page():
    if os.path.exists(RESULT_FILE):
        df = pd.read_excel(RESULT_FILE)
        results = df.to_dict(orient="records")
    else:
        results = []
    return render_template("results.html", results=results)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
