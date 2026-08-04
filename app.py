

import os, re, sys, warnings, tempfile
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import streamlit as st
warnings.filterwarnings("ignore")

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    import pdfplumber
    _PDF_OK = True
except ImportError:
    _PDF_OK = False

try:
    from docx import Document as _DocxDoc
    _DOCX_OK = True
except ImportError:
    _DOCX_OK = False


st.set_page_config(
    page_title="Skill Gap Analyser",
    page_icon="🎯",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
.main-title {
    font-size:2.4rem; font-weight:800;
    background:linear-gradient(135deg,#1a73e8,#9c27b0);
    -webkit-background-clip:text; -webkit-text-fill-color:transparent;
}
.subtitle { color:#555; font-size:1rem; margin-bottom:1rem; }
.sev-high   { background:#fde8e8; border-left:5px solid #e53935; padding:14px 18px; border-radius:8px; margin:8px 0; }
.sev-medium { background:#fff8e1; border-left:5px solid #fb8c00; padding:14px 18px; border-radius:8px; margin:8px 0; }
.sev-low    { background:#e8f5e9; border-left:5px solid #43a047; padding:14px 18px; border-radius:8px; margin:8px 0; }
.metric-box { background:#f0f4ff; border-radius:12px; padding:16px 10px; text-align:center; border:1px solid #c5cae9; }
.metric-val { font-size:2rem; font-weight:800; color:#1a237e; }
.metric-lbl { font-size:.78rem; color:#666; margin-top:4px; }
.skill-have    { display:inline-block; background:#d4edda; color:#155724; border-radius:14px; padding:3px 11px; margin:3px 2px; font-size:.82rem; font-weight:600; }
.skill-missing { display:inline-block; background:#f8d7da; color:#721c24; border-radius:14px; padding:3px 11px; margin:3px 2px; font-size:.82rem; font-weight:600; }
.course-card { background:#f8f9fa; border:1px solid #dee2e6; border-radius:10px; padding:12px 16px; margin:6px 0; }
.badge-free { background:#28a745; color:white; border-radius:4px; padding:2px 8px; font-size:.72rem; font-weight:700; }
.badge-paid { background:#6c757d; color:white; border-radius:4px; padding:2px 8px; font-size:.72rem; font-weight:700; }
.step-box { background:#f8f9fa; border-radius:10px; padding:12px 16px; margin:4px 0; border-left:4px solid #1a73e8; }
.profile-card { background:#f0f4ff; border-radius:10px; padding:14px; border:1px solid #c5cae9; margin-top:8px; }
</style>
""", unsafe_allow_html=True)


TECH_SKILLS = [
    "machine learning", "sql", "data analysis", "cloud computing",
    "deep learning", "natural language processing", "computer vision",
    "data engineering", "devops", "data visualization", "statistical analysis",
    "feature engineering", "model deployment", "api development",
    "cybersecurity", "networking", "mlops", "data pipelines", "etl",
    "big data", "blockchain", "database management",
]
PROG_LANGS = [
    "python", "java", "javascript", "r", "c++", "c#", "go", "scala",
    "kotlin", "swift", "typescript", "matlab", "bash", "ruby", "rust", "php",
]
SOFT_SKILLS = [
    "leadership", "communication", "teamwork", "adaptability",
    "time management", "critical thinking", "creativity", "collaboration",
    "presentation", "negotiation", "project management", "analytical thinking",
    "attention to detail", "problem solving",
]
TOOLS_KW = [
    "tensorflow", "pytorch", "scikit-learn", "pandas", "numpy", "tableau",
    "power bi", "excel", "spark", "hadoop", "docker", "kubernetes",
    "aws", "azure", "gcp", "git", "flask", "django", "fastapi",
    "mysql", "postgresql", "mongodb", "kafka", "airflow", "snowflake", "databricks",
]

CAREER_MAP = {
    "Data Scientist": [
        "data scientist",
    ],
    "AI Engineer": [
        "ai engineer", "artificial intelligence engineer", "machine learning engineer",
        "ml engineer", "nlp engineer", "deep learning engineer", "computer vision engineer",
        "generative ai engineer", "applied ai engineer",
    ],
    "Software Developer": [
        "software engineer", "software developer",
        "backend developer", "frontend developer", "full stack", "python developer",
    ],
    "Data Analyst": ["data analyst"],
    "Data Engineer": ["data engineer"],
    "Cloud / DevOps Engineer": ["cloud engineer", "devops", "cloud architect"],
    "Cybersecurity Expert": ["cybersecurity", "security analyst"],
    "Business Analyst": ["business analyst"],
}

COURSES = {
    "machine learning":            ("ML Specialization — Andrew Ng",      "Coursera",      True),
    "deep learning":               ("Practical Deep Learning",             "fast.ai",       True),
    "natural language processing": ("HuggingFace NLP Course",              "HuggingFace",   True),
    "computer vision":             ("CS231n — CNN for Vision",             "Stanford/YT",   True),
    "data analysis":               ("Kaggle Pandas Micro-course",          "Kaggle",        True),
    "data visualization":          ("Storytelling with Data",              "Book",          False),
    "statistical analysis":        ("StatQuest with Josh Starmer",         "YouTube",       True),
    "feature engineering":         ("Feature Engineering Micro-course",    "Kaggle",        True),
    "model deployment":            ("MLOps Specialization",                "DeepLearning.AI", False),
    "cloud computing":             ("AWS Cloud Practitioner Essentials",   "AWS",           True),
    "devops":                      ("Intro to DevOps",                     "edX/Linux Fdn", True),
    "cybersecurity":               ("Google Cybersecurity Certificate",    "Coursera",      False),
    "data engineering":            ("Data Engineering Zoomcamp",           "DataTalks",     True),
    "api development":             ("FastAPI Official Docs",               "FastAPI",       True),
    "sql":                         ("SQLZoo Interactive Tutorial",         "SQLZoo",        True),
    "data pipelines":              ("Data Engineering Zoomcamp",           "DataTalks",     True),
    "mlops":                       ("MLOps Specialization",                "DeepLearning.AI", False),
    "big data":                    ("Big Data Specialization",             "Coursera/UCSD", False),
    "networking":                  ("CompTIA Network+ Guide",              "CompTIA",       False),
    "etl":                         ("Data Engineering Zoomcamp",           "DataTalks",     True),
    "python":                      ("Real Python Tutorials",               "realpython.com", True),
    "r":                           ("R for Data Science",                  "r4ds.hadley.nz", True),
    "java":                        ("Java MOOC",                           "Univ Helsinki", True),
    "javascript":                  ("The Odin Project",                    "theodinproject.com", True),
    "scala":                       ("Rock the JVM",                        "RockTheJVM",    False),
    "bash":                        ("Linux Command Line Basics",           "LinuxCommand.org", True),
    "go":                          ("Official Go Tour",                    "go.dev",        True),
    "c++":                         ("learncpp.com",                        "learncpp.com",  True),
    "typescript":                  ("TypeScript Handbook",                 "typescriptlang.org", True),
    "leadership":                  ("Inspiring & Motivating Individuals",  "Coursera",      False),
    "communication":               ("Improving Communication Skills",      "Coursera",      False),
    "problem solving":             ("LeetCode Daily Challenges",           "LeetCode",      True),
    "project management":          ("Google PM Certificate",               "Coursera",      False),
    "critical thinking":           ("Critical Thinking & Problem Solving", "edX/RIT",       True),
    "teamwork":                    ("Leading Teams",                       "Coursera",      False),
    "time management":             ("Work Smarter Not Harder",             "Coursera",      True),
    "presentation":                ("Presentation Skills",                 "Coursera",      False),
    "adaptability":                ("Mindshift",                           "Coursera",      True),
    "analytical thinking":         ("Data-driven Decision Making",         "Coursera/PwC",  False),
}

DATA_FILE     = os.path.join(os.path.dirname(os.path.abspath(__file__)), "postings.csv")
PROFILES_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "career_profiles.json")

# SVM model artefacts
MODEL_DIR = os.path.dirname(os.path.abspath(__file__))

CAREER_TO_CAT = {
    "Data Scientist":        "Data Science",
    "AI Engineer":           "Data Science",
    "Software Developer":    "Software Engineering",
    "Data Analyst":          "Data Science",
    "Data Engineer":         "Data Engineering",
    "Cloud / DevOps Engineer": "Cloud & Infrastructure",
    "Cybersecurity Expert":  "Cybersecurity",
    "Business Analyst":      "Business & Management",
}

@st.cache_resource(show_spinner="Loading SVM model...")
def load_svm_model():
    import joblib
    paths = ["svm_model.pkl","mlb_tech.pkl","mlb_prog.pkl",
             "mlb_soft.pkl","le_y.pkl","le_cat.pkl",
             "feature_columns.pkl","salary_median.pkl"]
    if not all(os.path.exists(os.path.join(MODEL_DIR, p)) for p in paths):
        return None
    return {
        "model":    joblib.load(os.path.join(MODEL_DIR, "svm_model.pkl")),
        "mlb_tech": joblib.load(os.path.join(MODEL_DIR, "mlb_tech.pkl")),
        "mlb_prog": joblib.load(os.path.join(MODEL_DIR, "mlb_prog.pkl")),
        "mlb_soft": joblib.load(os.path.join(MODEL_DIR, "mlb_soft.pkl")),
        "le_y":     joblib.load(os.path.join(MODEL_DIR, "le_y.pkl")),
        "le_cat":   joblib.load(os.path.join(MODEL_DIR, "le_cat.pkl")),
        "cols":     joblib.load(os.path.join(MODEL_DIR, "feature_columns.pkl")),
        "sal_med":  joblib.load(os.path.join(MODEL_DIR, "salary_median.pkl")),
    }

def predict_severity_svm(artefacts, cv, selected_career):
    """Build feature vector from CV skills and predict severity using SVM."""
    try:
        tech_bin = artefacts["mlb_tech"].transform([cv["tech"]])
        prog_bin = artefacts["mlb_prog"].transform([cv["prog"]])
        soft_bin = artefacts["mlb_soft"].transform([cv["soft"]])

        cat_label = CAREER_TO_CAT.get(selected_career, "Data Science")
        known_cats = list(artefacts["le_cat"].classes_)
        if cat_label not in known_cats:
            cat_label = known_cats[0]
        cat_enc = artefacts["le_cat"].transform([cat_label])[0]

        exp_enc  = min(cv["tech_rating"] - 1, 5)
        n_tech   = len(cv["tech"])
        n_prog   = len(cv["prog"])
        n_soft   = len(cv["soft"])
        n_tools  = len(cv["tools"])
        total    = n_tech + n_prog + n_soft + n_tools

        base = np.array([[cat_enc, exp_enc, n_tech, n_prog,
                          n_soft, n_tools, total,
                          artefacts["sal_med"], 0]])

        tech_df = pd.DataFrame(tech_bin,
            columns=["tech_" + s.replace(" ","_") for s in artefacts["mlb_tech"].classes_])
        prog_df = pd.DataFrame(prog_bin,
            columns=["lang_" + p for p in artefacts["mlb_prog"].classes_])
        soft_df = pd.DataFrame(soft_bin,
            columns=["soft_" + s.replace(" ","_").replace(".","") for s in artefacts["mlb_soft"].classes_])

        base_df = pd.DataFrame(base, columns=[
            "cat_encoded","exp_encoded","tech_count","prog_count",
            "soft_count","tools_count","total_skills","salary_filled","is_remote"])

        row = pd.concat([base_df, tech_df, prog_df, soft_df], axis=1)

        # Align columns to training feature set
        for col in artefacts["cols"]:
            if col not in row.columns:
                row[col] = 0
        row = row[artefacts["cols"]]

        pred  = artefacts["model"].predict(row)[0]
        label = artefacts["le_y"].inverse_transform([pred])[0]
        return label
    except Exception:
        return None



@st.cache_resource(show_spinner="Loading career benchmarks...")
def build_career_profiles():
    import json
    # Load pre-computed profiles (used in deployment — no CSV needed)
    if os.path.exists(PROFILES_FILE):
        with open(PROFILES_FILE, "r") as f:
            return json.load(f)

    # Fallback: compute from raw CSV if available locally
    if not os.path.exists(DATA_FILE):
        st.error("career_profiles.json not found. Please add it to the project folder.")
        st.stop()

    df = pd.read_csv(DATA_FILE, low_memory=False)

    def map_cat(title):
        t = str(title).lower()
        for career, kws in CAREER_MAP.items():
            if any(kw in t for kw in kws):
                return career
        return None

    df["career"] = df["title"].apply(map_cat)
    df = df[df["career"].notna()].copy()

    def extract_skills_from_text(text):
        t = str(text).lower()
        tech  = [s for s in TECH_SKILLS  if s in t]
        prog  = [p for p in PROG_LANGS   if re.search(r"\b" + re.escape(p) + r"\b", t)]
        soft  = [s for s in SOFT_SKILLS  if s in t]
        tools = [tk for tk in TOOLS_KW   if tk in t]
        return tech, prog, soft, tools

    profiles = {}
    for career, grp in df.groupby("career"):
        all_tech, all_prog, all_soft, all_tools = [], [], [], []
        for desc in grp["description"].dropna():
            t, p, s, tk = extract_skills_from_text(desc)
            all_tech.extend(t); all_prog.extend(p)
            all_soft.extend(s); all_tools.extend(tk)

        n = len(grp)
        threshold = max(1, int(n * 0.15))
        from collections import Counter
        def top_skills(lst):
            return [s for s, c in Counter(lst).most_common() if c >= threshold]

        exp_map = {"Entry level": 1, "Associate": 2, "Mid-Senior level": 3,
                   "Director": 4, "Executive": 5, "Internship": 0}
        exp_scores = grp["formatted_experience_level"].map(exp_map).dropna()
        avg_exp = round(exp_scores.mean(), 1) if len(exp_scores) > 0 else 2.0

        sal = grp["normalized_salary"].dropna()
        avg_sal = int(sal.median()) if len(sal) > 0 else None

        profiles[career] = {
            "required_tech":  top_skills(all_tech),
            "required_prog":  top_skills(all_prog),
            "required_soft":  top_skills(all_soft),
            "required_tools": top_skills(all_tools)[:10],
            "avg_exp_level":  avg_exp,
            "avg_salary":     avg_sal,
            "total_postings": n,
        }

    return profiles



def extract_cv_text(file_obj):
    suffix = os.path.splitext(file_obj.name)[1].lower()
    if suffix == ".pdf":
        if not _PDF_OK:
            st.error("Install pdfplumber: pip install pdfplumber")
            return ""
        with pdfplumber.open(file_obj) as pdf:
            return "\n".join(p.extract_text() or "" for p in pdf.pages)
    elif suffix in (".docx", ".doc"):
        if not _DOCX_OK:
            st.error("Install python-docx: pip install python-docx")
            return ""
        doc = _DocxDoc(file_obj)
        return "\n".join(p.text for p in doc.paragraphs if p.text.strip())
    elif suffix == ".txt":
        raw = file_obj.read()
        return raw.decode("utf-8", errors="ignore") if isinstance(raw, bytes) else raw
    return ""


def parse_cv(text):
    tl = text.lower()
    tech  = [s for s in TECH_SKILLS  if s in tl]
    prog  = [p for p in PROG_LANGS   if re.search(r"\b" + re.escape(p) + r"\b", tl)]
    soft  = [s for s in SOFT_SKILLS  if s in tl]
    tools = [tk for tk in TOOLS_KW   if tk in tl]

    yrs = re.findall(r"(\d+)\+?\s*year", tl)
    exp = sum(int(y) for y in yrs) if yrs else 0

    senior_kw = ["senior", "lead", "principal", "architect", "director", "head of"]
    junior_kw = ["intern", "trainee", "fresher", "junior", "entry level"]

    if any(k in tl for k in senior_kw) or exp >= 5:   rating = 5
    elif exp >= 3 or len(tech) >= 6:                   rating = 4
    elif exp >= 1 or len(tech) >= 3:                   rating = 3
    elif any(k in tl for k in junior_kw):              rating = 2
    else:                                              rating = max(2, min(4, len(tech)))

    has_projects = bool(re.search(
        r"(project|built|developed|implemented|designed|deployed|created)", tl))

    return {
        "tech": tech, "prog": prog, "soft": soft, "tools": tools,
        "tech_rating": rating, "exp_years": exp, "has_projects": has_projects,
    }



def chart_donut(coverage):
    fig, ax = plt.subplots(figsize=(3.2, 3.2))
    ax.pie(
        [coverage, 100 - coverage],
        labels=[f"{coverage:.0f}%\nHave", f"{100-coverage:.0f}%\nGap"],
        colors=["#2ECC71", "#E74C3C"],
        startangle=90,
        wedgeprops={"width": 0.55},
        textprops={"fontsize": 11, "fontweight": "bold"},
    )
    ax.set_title("Skill Coverage", fontweight="bold", fontsize=11)
    plt.tight_layout()
    return fig


def chart_category(have_t, have_p, have_s, miss_t, miss_p, miss_s):
    fig, ax = plt.subplots(figsize=(4, 3.2))
    cats = ["Technical", "Languages", "Soft Skills"]
    have = [len(have_t), len(have_p), len(have_s)]
    miss = [len(miss_t), len(miss_p), len(miss_s)]
    x, w = np.arange(3), 0.35
    ax.bar(x - w/2, have, w, label="Have",    color="#2ECC71", edgecolor="black")
    ax.bar(x + w/2, miss, w, label="Missing", color="#E74C3C", edgecolor="black")
    ax.set_xticks(x); ax.set_xticklabels(cats, fontsize=9)
    ax.set_title("Skills by Category", fontweight="bold", fontsize=11)
    ax.set_ylabel("Count"); ax.legend(fontsize=8)
    plt.tight_layout()
    return fig


def chart_severity(severity):
    fig, ax = plt.subplots(figsize=(3.5, 3.2))
    levels = ["Low", "Medium", "High"]
    colors = ["#2ECC71", "#F39C12", "#E74C3C"]
    bars   = ax.bar(levels, [1, 1, 1], color=colors, edgecolor="black", width=0.5)
    for bar, level in zip(bars, levels):
        if level == severity:
            bar.set_linewidth(4)
            ax.text(bar.get_x() + bar.get_width() / 2,
                    bar.get_height() + 0.06,
                    "◀ YOU", ha="center", fontsize=10, fontweight="bold")
    ax.set_ylim(0, 1.5); ax.set_yticks([])
    ax.set_title("Severity Level", fontweight="bold", fontsize=11)
    plt.tight_layout()
    return fig


def chart_top_missing(all_missing):
    if not all_missing:
        return None
    top = all_missing[:8]
    fig, ax = plt.subplots(figsize=(5, max(3, len(top) * 0.5)))
    ax.barh(top[::-1], [1] * len(top), color="#E74C3C", edgecolor="black")
    ax.set_xlim(0, 1.4); ax.set_xticks([])
    ax.set_title("Top Skills to Acquire", fontweight="bold", fontsize=11)
    for i, s in enumerate(top[::-1]):
        ax.text(0.05, i, s.title(), va="center", fontsize=9, fontweight="600", color="white")
    plt.tight_layout()
    return fig



def main():
    st.markdown('<div class="main-title">🎯 Skill Gap Analyser</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="subtitle">Select your target career, upload your CV — '
        'get an instant skill gap report with course recommendations.</div>',
        unsafe_allow_html=True,
    )
    st.divider()


    if not os.path.exists(PROFILES_FILE) and not os.path.exists(DATA_FILE):
        st.error("career_profiles.json not found. Please add it to the project folder.")
        st.stop()

    profiles = build_career_profiles()
    career_options = sorted(profiles.keys())


    with st.sidebar:
        st.header("⚙️ Settings")

        selected_career = st.selectbox(
            "🎯 Target Career",
            career_options,
            index=0,
            help="Select the career you are targeting or hiring for.",
        )

        profile = profiles[selected_career]
        st.markdown("---")
        st.markdown(f"**📋 Benchmark: {selected_career}**")
        st.markdown(f"Based on **{profile['total_postings']} real job postings**")

        if profile["required_tech"]:
            st.markdown("**Required Tech Skills:**")
            for s in profile["required_tech"][:6]:
                st.markdown(f"  • {s.title()}")

        if profile["required_prog"]:
            st.markdown("**Required Languages:**")
            for p in profile["required_prog"][:5]:
                st.markdown(f"  • {p.title()}")

        if profile["avg_salary"]:
            st.markdown(f"**Avg Salary:** ${profile['avg_salary']:,}")

        st.markdown("---")
        st.markdown("**Supported formats:** PDF · DOCX · TXT")


    st.markdown("### 📄 Upload Your CV")
    st.caption("Your CV is analysed locally — not stored anywhere.")

    col1, col2 = st.columns([1, 1], gap="large")
    with col1:
        cv_file = st.file_uploader(
            "Upload CV (PDF / DOCX / TXT)",
            type=["pdf", "docx", "doc", "txt"],
            label_visibility="collapsed",
        )
    with col2:
        cv_paste = st.text_area(
            "Or paste your CV text",
            height=200,
            placeholder="Paste your resume / CV content here...",
            label_visibility="collapsed",
        )

    st.divider()
    run = st.button("🔍 Analyse My Skill Gap", type="primary", use_container_width=True)

    if run:

        cv_text = ""
        if cv_file:
            cv_file.seek(0)
            cv_text = extract_cv_text(cv_file)
        elif cv_paste.strip():
            cv_text = cv_paste.strip()
        else:
            st.warning("⚠️ Please upload a CV or paste your CV text above.")
            st.stop()

        if not cv_text.strip():
            st.error("Could not extract text from the file. Try pasting the text directly.")
            st.stop()


        with st.spinner("Extracting skills from your CV..."):
            cv = parse_cv(cv_text)

        profile = profiles[selected_career]


        need_tech = set(profile["required_tech"])
        need_prog = set(profile["required_prog"])
        need_soft = set(profile["required_soft"])
        have_tech = set(cv["tech"])
        have_prog = set(cv["prog"])
        have_soft = set(cv["soft"])

        miss_tech = sorted(need_tech - have_tech)
        miss_prog = sorted(need_prog - have_prog)
        miss_soft = sorted(need_soft - have_soft)

        matched_tech = sorted(have_tech & need_tech)
        matched_prog = sorted(have_prog & need_prog)
        matched_soft = sorted(have_soft & need_soft)

        need_all   = need_tech | need_prog | need_soft
        have_all   = have_tech | have_prog | have_soft
        total_req  = len(need_all)
        total_have = len(have_all & need_all)
        coverage   = (total_have / total_req * 100) if total_req > 0 else 100.0
        all_missing = miss_tech + miss_prog + miss_soft


        # Use trained SVM model if available, else fall back to rule-based
        svm_artefacts = load_svm_model()
        svm_severity  = predict_severity_svm(svm_artefacts, cv, selected_career) if svm_artefacts else None

        if svm_severity:
            severity = svm_severity
        elif coverage >= 75:
            severity = "Low"
        elif coverage >= 45:
            severity = "Medium"
        else:
            severity = "High"


        st.divider()
        st.header(f"📊 Results — {selected_career}")
        if svm_severity:
            st.caption("🤖 Severity predicted by trained **SVM (RBF)** model")
        else:
            st.caption("📐 Severity calculated by rule-based coverage threshold")


        sev_map = {
            "Low":    ("sev-low",    "✅", "Strong match — minor gaps only. You're nearly job-ready!"),
            "Medium": ("sev-medium", "⚠️", "Moderate gap — targeted upskilling will make you competitive."),
            "High":   ("sev-high",   "❌", "Significant gap — a focused learning plan is recommended."),
        }
        cls_name, icon, msg = sev_map[severity]
        st.markdown(
            f'<div class="{cls_name}">'
            f'<strong>{icon} Skill Gap Severity: {severity.upper()}</strong><br>{msg}'
            f'</div>',
            unsafe_allow_html=True,
        )

        # Metrics
        m1, m2, m3, m4 = st.columns(4)
        for col, val, lbl in [
            (m1, f"{coverage:.0f}%",          "Skill Coverage"),
            (m2, f"{total_have}/{total_req}",  "Skills Matched"),
            (m3, f"{len(all_missing)}",        "Skills to Acquire"),
            (m4, f"{cv['tech_rating']}/5",     "Your Tech Rating"),
        ]:
            col.markdown(
                f'<div class="metric-box">'
                f'<div class="metric-val">{val}</div>'
                f'<div class="metric-lbl">{lbl}</div>'
                f'</div>',
                unsafe_allow_html=True,
            )

        st.divider()


        ch1, ch2, ch3, ch4 = st.columns(4)
        ch1.pyplot(chart_donut(coverage), use_container_width=True)
        ch2.pyplot(chart_category(
            matched_tech, matched_prog, matched_soft,
            miss_tech, miss_prog, miss_soft),
            use_container_width=True)
        ch3.pyplot(chart_severity(severity), use_container_width=True)
        fig_miss = chart_top_missing(all_missing)
        if fig_miss:
            ch4.pyplot(fig_miss, use_container_width=True)
        plt.close("all")

        st.divider()


        col_have, col_miss = st.columns(2, gap="large")

        with col_have:
            st.subheader("✅ Skills You Already Have")
            if matched_tech:
                st.markdown("**Technical Skills**")
                st.markdown(
                    " ".join(f'<span class="skill-have">{s.title()}</span>'
                             for s in matched_tech),
                    unsafe_allow_html=True,
                )
            if matched_prog:
                st.markdown("**Programming Languages**")
                st.markdown(
                    " ".join(f'<span class="skill-have">{p.title()}</span>'
                             for p in matched_prog),
                    unsafe_allow_html=True,
                )
            if matched_soft:
                st.markdown("**Soft Skills**")
                st.markdown(
                    " ".join(f'<span class="skill-have">{s.title()}</span>'
                             for s in matched_soft),
                    unsafe_allow_html=True,
                )
            # Extra skills not in benchmark
            extra = sorted((have_tech | have_prog | have_soft) - need_all)
            if extra:
                st.markdown("**Bonus Skills (not required but useful)**")
                st.markdown(
                    " ".join(f'<span class="skill-have" style="background:#cce5ff;color:#004085">'
                             f'{s.title()}</span>' for s in extra),
                    unsafe_allow_html=True,
                )
            if cv["tools"]:
                st.markdown("**Tools Detected**")
                st.markdown(
                    " ".join(f'<span class="skill-have">{t}</span>'
                             for t in cv["tools"][:12]),
                    unsafe_allow_html=True,
                )

        with col_miss:
            st.subheader("❌ Skills You Still Need")
            if not all_missing:
                st.success("🎉 You meet all the benchmark requirements!")
            else:
                if miss_tech:
                    st.markdown("**Technical Skills**")
                    st.markdown(
                        " ".join(f'<span class="skill-missing">{s.title()}</span>'
                                 for s in miss_tech),
                        unsafe_allow_html=True,
                    )
                if miss_prog:
                    st.markdown("**Programming Languages**")
                    st.markdown(
                        " ".join(f'<span class="skill-missing">{p.title()}</span>'
                                 for p in miss_prog),
                        unsafe_allow_html=True,
                    )
                if miss_soft:
                    st.markdown("**Soft Skills**")
                    st.markdown(
                        " ".join(f'<span class="skill-missing">{s.title()}</span>'
                                 for s in miss_soft),
                        unsafe_allow_html=True,
                    )


        if all_missing:
            st.divider()
            st.subheader("📚 Course Recommendations")
            timeline = {
                "High":   "10–14 weeks intensive",
                "Medium": "4–8 weeks targeted",
                "Low":    "2–4 weeks polish",
            }
            st.info(f"⏱️ Suggested learning timeline: **{timeline[severity]}**")

            for skill in all_missing[:10]:
                course = COURSES.get(skill.lower())
                if course:
                    title, platform, free = course
                else:
                    title    = f"Search: {skill.title()} course"
                    platform = "Coursera / Udemy / YouTube"
                    free     = True

                badge = '<span class="badge-free">FREE</span>' if free else '<span class="badge-paid">PAID</span>'
                st.markdown(
                    f'<div class="course-card">'
                    f'<strong>🔧 {skill.title()}</strong><br>'
                    f'📖 {title}&nbsp;&nbsp;{badge}'
                    f'<span style="color:#777;font-size:.82rem">&nbsp;— {platform}</span>'
                    f'</div>',
                    unsafe_allow_html=True,
                )


        st.divider()
        st.subheader("🗺️ Your Personalised Action Plan")

        plans = {
            "High": [
                f"Focus on the top 3 missing skills: {', '.join(all_missing[:3]) or 'see above'}",
                "Build 2–3 portfolio projects that demonstrate the missing skills",
                "Contribute to open-source repositories on GitHub",
                "Apply for internships or junior roles to gain practical experience",
                "Update your LinkedIn and GitHub with new skills as you learn them",
            ],
            "Medium": [
                f"Complete one course for each missing skill (start with: {', '.join(all_missing[:2]) or 'see above'})",
                "Build a capstone project that covers your target skills",
                "Attempt weekly Kaggle challenges to practise data skills",
                "Network on LinkedIn with professionals in the target role",
                "Prepare for technical interviews using LeetCode / HackerRank",
            ],
            "Low": [
                "Quantify your CV achievements with metrics (e.g. 'improved X by Y%')",
                "Start applying broadly — your profile closely matches the benchmark",
                "Deepen expertise in 1–2 top skills to differentiate yourself",
                "Polish your GitHub portfolio with clean, well-documented projects",
                "Prepare mock interviews — you're nearly job-ready",
            ],
        }

        for i, step in enumerate(plans[severity], 1):
            st.markdown(
                f'<div class="step-box"><strong>{i}.</strong> {step}</div>',
                unsafe_allow_html=True,
            )

        if not cv["has_projects"]:
            st.warning("⚠️ No projects detected in your CV — adding a personal project will significantly strengthen your profile.")
        if cv["exp_years"] > 0:
            st.info(f"📅 Experience detected: approximately **{cv['exp_years']} year(s)**")


        st.divider()
        with st.expander("📄 Full Text Summary"):
            lines = [
                "=" * 60,
                "     SKILL GAP ANALYSIS REPORT",
                "=" * 60,
                f"  Target Career  : {selected_career}",
                f"  Benchmark      : {profile['total_postings']} real job postings",
                f"  Experience     : {cv['exp_years']} years detected",
                f"  Tech Rating    : {cv['tech_rating']} / 5",
                f"  Skill Coverage : {coverage:.1f}%  ({total_have} of {total_req} required)",
                f"  Severity       : {severity.upper()}",
                "-" * 60,
                "  SKILLS YOU HAVE (matching benchmark)",
                "  Tech       : " + (", ".join(matched_tech) or "None"),
                "  Languages  : " + (", ".join(matched_prog) or "None"),
                "  Soft Skills: " + (", ".join(matched_soft) or "None"),
                "-" * 60,
                "  SKILLS TO ACQUIRE",
                "  Tech       : " + (", ".join(miss_tech) or "None"),
                "  Languages  : " + (", ".join(miss_prog) or "None"),
                "  Soft Skills: " + (", ".join(miss_soft) or "None"),
                "=" * 60,
            ]
            st.code("\n".join(lines), language=None)


    st.divider()
    st.markdown(
        "<div style='text-align:center;color:#888;font-size:.8rem'>"
        "Skill Gap Analyser · Srikaran_Sankar · 2026"
        "<br>© 2026 Srikaran Sankar. All rights reserved."
        "</div>",
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
