

import os, re, sys, json, time, random, warnings
from datetime import datetime
import pandas as pd
import streamlit as st
import plotly.graph_objects as go
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

from openai import OpenAI


st.set_page_config(
    page_title="Skill Gap Analyser — Agent Edition",
    page_icon="🤖",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------------------------------------------------------------------------
# Validated palette (see dataviz skill references/palette.md) — status colors
# are fixed/reserved, categorical hues assigned in fixed order, never cycled.
# ---------------------------------------------------------------------------
COLOR_GOOD      = "#0ca30c"   # severity: Low / "have" skills
COLOR_WARNING   = "#fab219"   # severity: Medium
COLOR_CRITICAL  = "#d03b3b"   # severity: High / "missing" skills
COLOR_CAT_1     = "#2a78d6"   # categorical slot 1 — blue (Technical)
COLOR_CAT_2     = "#1baf7a"   # categorical slot 2 — aqua (Languages)
COLOR_CAT_3     = "#eda100"   # categorical slot 3 — yellow (Soft Skills)
COLOR_MUTED     = "#c3c2b7"   # muted ink / gap fill

st.markdown(f"""
<style>
.main-title {{
    font-size:2.5rem; font-weight:800;
    background:linear-gradient(135deg,#1a73e8,#9c27b0);
    -webkit-background-clip:text; -webkit-text-fill-color:transparent;
}}
.subtitle {{ color:#666; font-size:1.02rem; margin-bottom:.5rem; }}

.pulse-dot {{
    display:inline-block; width:9px; height:9px; border-radius:50%;
    background:{COLOR_GOOD}; margin-right:6px;
    box-shadow:0 0 0 rgba(12,163,12,.6); animation:pulse 2s infinite;
}}
.pulse-dot.off {{ background:#999; animation:none; box-shadow:none; }}
@keyframes pulse {{
    0%   {{ box-shadow:0 0 0 0 rgba(12,163,12,.5); }}
    70%  {{ box-shadow:0 0 0 8px rgba(12,163,12,0); }}
    100% {{ box-shadow:0 0 0 0 rgba(12,163,12,0); }}
}}

.card {{
    background:var(--card-bg,#fff); border:1px solid var(--card-border,#eceae4);
    border-radius:14px; padding:18px 20px; box-shadow:0 1px 3px rgba(0,0,0,.06);
    margin-bottom:10px;
}}
@media (prefers-color-scheme: dark) {{
    .card {{ --card-bg:#1c1c1a; --card-border:#33322e; box-shadow:0 1px 3px rgba(0,0,0,.4); }}
}}

.sev-banner {{ border-radius:14px; padding:18px 22px; margin:8px 0; border-left:6px solid; }}
.sev-High   {{ background:rgba(208,59,59,.08);  border-color:{COLOR_CRITICAL}; }}
.sev-Medium {{ background:rgba(250,178,25,.10); border-color:{COLOR_WARNING}; }}
.sev-Low    {{ background:rgba(12,163,12,.08);  border-color:{COLOR_GOOD}; }}

.metric-card {{
    background:var(--card-bg,#fff); border:1px solid var(--card-border,#eceae4);
    border-radius:12px; padding:14px 10px; text-align:center;
    box-shadow:0 1px 3px rgba(0,0,0,.05);
}}
.metric-val {{ font-size:1.9rem; font-weight:800; color:var(--ink-primary,#0b0b0b); }}
.metric-lbl {{ font-size:.76rem; color:var(--ink-secondary,#666); margin-top:2px; text-transform:uppercase; letter-spacing:.04em; }}
@media (prefers-color-scheme: dark) {{
    .metric-card {{ --card-bg:#1c1c1a; --card-border:#33322e; }}
    .metric-val {{ --ink-primary:#fff; }}
    .metric-lbl {{ --ink-secondary:#c3c2b7; }}
}}

.skill-have    {{ display:inline-block; background:rgba(12,163,12,.12); color:{COLOR_GOOD}; border:1px solid rgba(12,163,12,.3); border-radius:14px; padding:4px 12px; margin:3px 3px; font-size:.82rem; font-weight:600; }}
.skill-missing {{ display:inline-block; background:rgba(208,59,59,.10); color:{COLOR_CRITICAL}; border:1px solid rgba(208,59,59,.3); border-radius:14px; padding:4px 12px; margin:3px 3px; font-size:.82rem; font-weight:600; }}
.skill-bonus   {{ display:inline-block; background:rgba(42,120,214,.10); color:{COLOR_CAT_1}; border:1px solid rgba(42,120,214,.3); border-radius:14px; padding:4px 12px; margin:3px 3px; font-size:.82rem; font-weight:600; }}

.priority-chip {{
    display:flex; align-items:center; gap:10px; background:var(--card-bg,#fff);
    border:1px solid var(--card-border,#eceae4); border-left:4px solid {COLOR_CRITICAL};
    border-radius:10px; padding:9px 14px; margin:5px 0;
}}
.priority-num {{ font-weight:800; color:{COLOR_CRITICAL}; font-size:.95rem; min-width:20px; }}
@media (prefers-color-scheme: dark) {{ .priority-chip {{ --card-bg:#1c1c1a; --card-border:#33322e; }} }}

.course-card {{ background:var(--card-bg,#fff); border:1px solid var(--card-border,#eceae4); border-radius:10px; padding:12px 16px; margin:6px 0; }}
.badge-free {{ background:{COLOR_GOOD}; color:white; border-radius:4px; padding:2px 8px; font-size:.72rem; font-weight:700; }}
.badge-paid {{ background:#6c757d; color:white; border-radius:4px; padding:2px 8px; font-size:.72rem; font-weight:700; }}
@media (prefers-color-scheme: dark) {{ .course-card {{ --card-bg:#1c1c1a; --card-border:#33322e; }} }}

.step-box {{ background:var(--card-bg,#fff); border-radius:10px; padding:12px 16px; margin:5px 0; border-left:4px solid {COLOR_CAT_1}; }}
@media (prefers-color-scheme: dark) {{ .step-box {{ --card-bg:#1c1c1a; }} }}

.reasoning-box {{ background:rgba(74,58,167,.07); border-left:5px solid #4a3aa7; padding:16px 20px; border-radius:12px; margin:10px 0; font-style:italic; line-height:1.55; }}

.hero-wrap {{ text-align:center; padding:4px 0 2px; }}
.hero-badge {{
    display:inline-block; background:rgba(42,120,214,.10); color:{COLOR_CAT_1};
    border:1px solid rgba(42,120,214,.25); border-radius:20px; padding:4px 14px;
    font-size:.75rem; font-weight:700; letter-spacing:.04em; margin-bottom:12px;
    text-transform:uppercase;
}}
.stat-row {{ display:flex; gap:14px; justify-content:center; flex-wrap:wrap; margin:20px 0 4px; }}
.stat-chip {{
    background:var(--card-bg,#fff); border:1px solid var(--card-border,#eceae4);
    border-radius:14px; padding:10px 24px; text-align:center; min-width:130px;
    box-shadow:0 1px 3px rgba(0,0,0,.05);
}}
.stat-chip .num {{ font-size:1.5rem; font-weight:800; color:{COLOR_CAT_1}; display:block; }}
.stat-chip .lbl {{ font-size:.72rem; color:var(--ink-secondary,#888); text-transform:uppercase; letter-spacing:.03em; }}

.steps-row {{ display:flex; gap:16px; margin:6px 0 4px; }}
.step-card2 {{
    flex:1; background:var(--card-bg,#fff); border:1px solid var(--card-border,#eceae4);
    border-radius:14px; padding:18px 16px; text-align:center; box-shadow:0 1px 3px rgba(0,0,0,.05);
}}
.step-num2 {{
    display:inline-block; width:24px; height:24px; border-radius:50%; background:{COLOR_CAT_1};
    color:#fff; font-size:.74rem; font-weight:800; line-height:24px; margin-bottom:8px;
}}
.step-card2 .step-icon {{ font-size:1.5rem; display:block; margin-bottom:2px; }}
.step-card2 h4 {{ margin:4px 0 4px; font-size:.9rem; }}
.step-card2 p {{ margin:0; font-size:.78rem; color:var(--ink-secondary,#888); }}

@media (prefers-color-scheme: dark) {{
    .stat-chip, .step-card2 {{ --card-bg:#1c1c1a; --card-border:#33322e; --ink-secondary:#c3c2b7; }}
}}
</style>
""", unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Skill taxonomy — identical to skill_gap_analysis.ipynb / agent.ipynb so the
# agent's severity thresholds stay consistent with how the labels were built.
# ---------------------------------------------------------------------------
TECH_SKILLS = [
    "machine learning", "sql", "data analysis", "cloud computing",
    "deep learning", "natural language processing", "computer vision",
    "data engineering", "devops", "data visualization", "statistical analysis",
    "feature engineering", "model deployment", "api development",
    "cybersecurity", "networking", "mlops", "data pipelines", "etl",
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
    "power bi", "powerbi", "excel", "spark", "hadoop", "docker", "kubernetes",
    "aws", "azure", "gcp", "git", "flask", "django", "fastapi",
    "mysql", "postgresql", "mongodb", "kafka", "airflow", "snowflake", "databricks",
]

# Tool names / abbreviations / common phrasings that count as evidence for a
# taxonomy skill even when the literal phrase isn't present (e.g. a CV says
# "Tableau, PowerBI" rather than the literal words "data visualization").
# Deterministic — no LLM judgment involved, just widening what counts as a
# keyword match. Covers all TECH_SKILLS, PROG_LANGS, and SOFT_SKILLS taxonomy
# terms with well-established real-world synonyms, not speculative guesses.
# Matched against text with hyphens normalized to spaces (see extract_skills_from_text),
# so aliases are written without hyphens (e.g. "detail oriented" not "detail-oriented").
SKILL_ALIASES = {
    # --- TECH_SKILLS ---
    "machine learning":            ["scikit-learn", "tensorflow", "pytorch", "keras", "xgboost",
                                     "lightgbm", "random forest", "ml model", "supervised learning",
                                     "unsupervised learning"],
    "data analysis":                ["pandas", "data analytics", "exploratory data analysis", "eda"],
    "cloud computing":              ["aws", "azure", "gcp", "google cloud", "amazon web services",
                                     "cloud platform", "cloud infrastructure"],
    "deep learning":                ["tensorflow", "pytorch", "keras", "neural network", "cnn", "rnn", "transformer"],
    "natural language processing": ["nlp", "llm", "llms", "large language model", "generative ai",
                                     "gpt", "bert", "chatbot", "text mining"],
    "computer vision":              ["opencv", "image recognition", "object detection", "image processing"],
    "data engineering":              ["spark", "hadoop", "databricks", "data lake", "data lakes",
                                     "data warehouse", "big data"],
    "devops":                        ["docker", "kubernetes", "ci/cd", "jenkins", "terraform", "ansible"],
    "data visualization":           ["tableau", "power bi", "powerbi", "looker", "qlik",
                                     "matplotlib", "seaborn", "dashboards", "data viz"],
    "statistical analysis":         ["hypothesis testing", "regression analysis", "anova", "statistics",
                                     "statistical modeling"],
    "feature engineering":          ["feature selection", "feature extraction"],
    "model deployment":             ["streamlit", "docker", "kubernetes", "mlflow", "model serving",
                                     "production deployment"],
    "api development":              ["fastapi", "flask", "django", "rest api", "restful", "graphql"],
    "cybersecurity":                 ["penetration testing", "security audit", "information security",
                                     "infosec", "threat detection"],
    "networking":                    ["tcp/ip", "vpn", "network security", "firewall", "dns"],
    "mlops":                          ["mlflow", "kubeflow", "model monitoring", "model versioning"],
    "data pipelines":                ["airflow", "kafka", "elt pipeline", "elt", "data flow"],
    "etl":                            ["elt", "elt pipeline", "extract transform load", "data integration"],
    # --- SOFT_SKILLS ---
    "leadership":                    ["team lead", "led a team", "managed a team"],
    "teamwork":                      ["team player", "cross functional team", "collaborative team"],
    "critical thinking":             ["critical analysis", "problem analysis"],
    "creativity":                    ["creative thinking", "innovative"],
    "collaboration":                 ["cross functional collaboration", "team collaboration"],
    "presentation":                  ["public speaking", "stakeholder presentation"],
    "project management":           ["pmp", "agile", "scrum", "project lead", "program management"],
    "analytical thinking":          ["analytical skills", "data driven decision making"],
    "attention to detail":          ["detail oriented", "meticulous"],
    "problem solving":               ["problem solving skills"],  # hyphen variant handled by normalization
    "time management":               ["prioritization", "deadline management"],
    "adaptability":                  ["flexibility", "adaptable"],
}

# Same idea as SKILL_ALIASES but for PROG_LANGS, matched with word-boundary
# safety (has_word) since these tokens can be short/ambiguous.
PROG_ALIASES = {
    "javascript": ["js", "node.js", "nodejs", "node"],
    "typescript": ["ts"],
    "go":          ["golang"],
    "c#":          [".net", "dotnet", "c sharp"],
}

TECH_TITLES = [
    'data scientist','data analyst','machine learning','ml engineer',
    'software engineer','software developer','data engineer',
    'ai engineer','cloud engineer','devops','cloud architect',
    'cybersecurity','security analyst','business analyst',
    'backend developer','frontend developer','full stack',
    'python developer','nlp engineer','computer vision','deep learning'
]

CAREER_MAP = {
    "Data Scientist": ["data scientist"],
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

# 8-way UI career -> 7-way job_category used by the agent's benchmark thresholds
# (must mirror map_category() in skill_gap_analysis.ipynb / agent.ipynb exactly).
CAREER_TO_JOB_CATEGORY = {
    "Data Scientist":          "Data Scientist / ML",
    "AI Engineer":              "Data Scientist / ML",
    "Software Developer":       "Software Developer",
    "Data Analyst":              "Data Analyst",
    "Data Engineer":             "Data Engineer",
    "Cloud / DevOps Engineer":  "Cloud / DevOps",
    "Cybersecurity Expert":     "Cybersecurity",
    "Business Analyst":         "Business Analyst",
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

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_CANDIDATES = [
    os.path.join(SCRIPT_DIR, "postings.csv"),
]
TENSORX_BASE_URL = "https://api.tensorx.ai/v1"
MODEL = "deepseek/deepseek-v4-flash"
TRUNC_LEN = 3400  # bounded extraction window — mirrors skill_gap_agent_analysis.ipynb


# ---------------------------------------------------------------------------
# Data / benchmark loading (career skill profiles + agent threshold lookup)
# ---------------------------------------------------------------------------
def has_word(word, text):
    return bool(re.search(r"\b" + re.escape(word) + r"\b", text))


def extract_skills_from_text(text):
    t = str(text).lower().replace("-", " ")
    tech  = [s for s in TECH_SKILLS
             if s in t or any(alias in t for alias in SKILL_ALIASES.get(s, []))]
    prog  = [p for p in PROG_LANGS
             if has_word(p, t) or any(has_word(a, t) for a in PROG_ALIASES.get(p, []))]
    soft  = [s for s in SOFT_SKILLS
             if s in t or any(alias in t for alias in SKILL_ALIASES.get(s, []))]
    tools = [tk for tk in TOOLS_KW   if tk in t]
    return tech, prog, soft, tools


@st.cache_resource(show_spinner="Loading career benchmarks from LinkedIn postings...")
def load_benchmarks():
    """Single pass over postings.csv building both:
    - career_profiles: required skills per UI career (8-way), for the CV comparison UI
    - category_thresholds: q33/q66 total_skills per job_category (7-way), for the agent tool
    """
    data_file = next((p for p in DATA_CANDIDATES if os.path.exists(p)), None)
    if data_file is None:
        st.error(
            "postings.csv not found. Place it next to agent_app.py, or edit "
            "DATA_CANDIDATES at the top of the script to point at its location."
        )
        st.stop()

    df = pd.read_csv(data_file, low_memory=False)
    mask = df["title"].str.lower().apply(lambda x: any(t in str(x) for t in TECH_TITLES))
    df = df[mask].copy().reset_index(drop=True)

    def map_career(title):
        t = str(title).lower()
        for career, kws in CAREER_MAP.items():
            if any(kw in t for kw in kws):
                return career
        return None

    def map_job_category(title):
        t = str(title).lower()
        if any(x in t for x in ['data scientist','machine learning','ml engineer',
                                 'nlp engineer','deep learning','ai engineer','computer vision']):
            return 'Data Scientist / ML'
        elif any(x in t for x in ['software engineer','software developer',
                                   'backend','frontend','full stack','python developer']):
            return 'Software Developer'
        elif 'data analyst' in t:
            return 'Data Analyst'
        elif 'data engineer' in t:
            return 'Data Engineer'
        elif any(x in t for x in ['cloud','devops']):
            return 'Cloud / DevOps'
        elif any(x in t for x in ['cybersecurity','security analyst']):
            return 'Cybersecurity'
        elif 'business analyst' in t:
            return 'Business Analyst'
        else:
            return 'Other Tech'

    df["career"] = df["title"].apply(map_career)
    df["job_category"] = df["title"].apply(map_job_category)

    extracted = df["description"].apply(extract_skills_from_text)
    df["tech_skills"] = extracted.apply(lambda x: x[0])
    df["prog_langs"]  = extracted.apply(lambda x: x[1])
    df["soft_skills"] = extracted.apply(lambda x: x[2])
    df["tools"]       = extracted.apply(lambda x: x[3])
    # Soft skills excluded from total_skills: severity/coverage should track tech
    # readiness, not be inflated by soft-skill buzzwords the CV happens to contain.
    df["total_skills"] = df["tech_skills"].apply(len) + df["prog_langs"].apply(len)

    # --- career_profiles (8-way, for CV skill comparison / dashboard) ---
    career_profiles = {}
    for career, grp in df[df["career"].notna()].groupby("career"):
        n = len(grp)
        threshold = max(1, int(n * 0.15))
        from collections import Counter
        def top_skills(lst):
            flat = [item for sub in lst for item in sub]
            return [s for s, c in Counter(flat).most_common() if c >= threshold]

        exp_map = {"Entry level": 1, "Associate": 2, "Mid-Senior level": 3,
                   "Director": 4, "Executive": 5, "Internship": 0}
        exp_scores = grp["formatted_experience_level"].map(exp_map).dropna()
        avg_exp = round(exp_scores.mean(), 1) if len(exp_scores) > 0 else 2.0
        sal = grp["normalized_salary"].dropna()
        avg_sal = int(sal.median()) if len(sal) > 0 else None

        career_profiles[career] = {
            "required_tech":  top_skills(grp["tech_skills"]),
            "required_prog":  top_skills(grp["prog_langs"]),
            "required_soft":  top_skills(grp["soft_skills"]),
            "required_tools": top_skills(grp["tools"])[:10],
            "avg_exp_level":  avg_exp,
            "avg_salary":     avg_sal,
            "total_postings": n,
        }

    # --- category_thresholds (7-way, for the agent's get_category_thresholds tool) ---
    thresholds = df.groupby("job_category")["total_skills"].quantile([0.33, 0.66]).unstack()
    thresholds.columns = ["q33", "q66"]
    global_q33 = df["total_skills"].quantile(0.33)
    global_q66 = df["total_skills"].quantile(0.66)

    category_thresholds = {
        cat: {"q33": float(row["q33"]), "q66": float(row["q66"])}
        for cat, row in thresholds.iterrows()
    }
    category_thresholds["__global__"] = {"q33": float(global_q33), "q66": float(global_q66)}

    return career_profiles, category_thresholds


# ---------------------------------------------------------------------------
# DeepSeek agent (via TensorX) — mirrors agent.ipynb's tool-calling loop,
# adapted to classify a CV's skill-gap severity against a target career, with
# an on_step callback so the caller can stream live progress to the UI.
# ---------------------------------------------------------------------------
@st.cache_resource(show_spinner=False)
def get_client(api_key):
    return OpenAI(api_key=api_key, base_url=TENSORX_BASE_URL)


def get_category_thresholds_tool(category_thresholds, job_category):
    t = category_thresholds.get(job_category, category_thresholds["__global__"])
    return {
        "job_category": job_category,
        "q33_threshold": round(t["q33"], 2),
        "q66_threshold": round(t["q66"], 2),
        "rule": "severity = High if total_skills <= q33_threshold (candidate has fewer skills "
                "than most job postings in this category require); Low if total_skills > "
                "q66_threshold (candidate exceeds most postings' skill count); Medium otherwise.",
    }


AGENT_TOOLS = [
    {
        "type": "function",
        "function": {
            "name": "get_category_thresholds",
            "description": (
                "Look up the skill-count benchmark thresholds (33rd and 66th percentile of "
                "total_skills) for a given job category, computed from real LinkedIn postings. "
                "Always call this before deciding the final severity label."
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "job_category": {
                        "type": "string",
                        "description": "One of: Data Scientist / ML, Software Developer, "
                                        "Data Analyst, Data Engineer, Cloud / DevOps, "
                                        "Cybersecurity, Business Analyst, Other Tech",
                    }
                },
                "required": ["job_category"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "extract_skills_exact",
            "description": (
                "Runs the authoritative deterministic keyword-matching extraction on the raw CV text "
                "(bounded to the first 3400 characters) and returns the exact skill lists, "
                "total_skill_count, and computed severity. This is the single source of truth for the "
                "count — call this after get_category_thresholds, before submit_classification."
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "job_category": {"type": "string",
                        "description": "The candidate's target job category."},
                },
                "required": ["job_category"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "submit_classification",
            "description": "Submit your final skill-gap severity classification for this candidate's CV.",
            "parameters": {
                "type": "object",
                "properties": {
                    "extracted_technical_skills": {"type": "array", "items": {"type": "string"},
                        "description": "The technical skills list returned by extract_skills_exact."},
                    "extracted_programming_languages": {"type": "array", "items": {"type": "string"}},
                    "extracted_soft_skills": {"type": "array", "items": {"type": "string"}},
                    "total_skill_count": {"type": "integer",
                        "description": "The total_skill_count returned by extract_skills_exact — do not recompute or override it."},
                    "severity": {"type": "string", "enum": ["Low", "Medium", "High"],
                        "description": "Must exactly match the computed_severity returned by extract_skills_exact."},
                    "confidence": {"type": "number",
                        "description": "Self-assessed confidence in this label, 0.0-1.0."},
                    "reasoning": {"type": "string",
                        "description": "2-4 sentence justification referencing the thresholds and extracted skills, written for the candidate to read."},
                },
                "required": ["severity", "total_skill_count", "reasoning"],
            },
        },
    },
]


def build_system_prompt():
    return f"""You are a career-advisory agent that analyses a candidate's CV/resume against a \
target job category and classifies their skill-gap severity as Low, Medium, or High.

You have access to a deterministic skill-extraction tool that runs the authoritative keyword-matching \
algorithm used to build the benchmark thresholds themselves. Do NOT attempt to identify or count \
skills yourself from the CV text — extract_skills_exact's output is the single source of truth for \
total_skill_count, exactly as it was for building the category benchmarks from real job postings. \
Note the tool only reads the first {TRUNC_LEN} characters of the CV (a bounded processing window, \
consistent with real context-length constraints).

Skill taxonomy (for your own understanding — the extraction tool applies this automatically):
Technical skills: {', '.join(TECH_SKILLS)}
Programming languages: {', '.join(PROG_LANGS)}
Soft skills: {', '.join(SOFT_SKILLS)}

Workflow (mandatory, in order):
1. Call get_category_thresholds with the candidate's target job_category.
2. Call extract_skills_exact with the same job_category — trust its result completely.
3. Call submit_classification with the severity and skill lists EXACTLY as returned by \
   extract_skills_exact (do not override or second-guess them), plus a confidence score and a short \
   reasoning written directly to the candidate (e.g. "Your CV shows...") that explains the result in \
   plain language.

Always finish by calling submit_classification — that is the only way to record your answer."""


def extract_skills_exact_tool(category_thresholds, job_category, cv_text):
    """Deterministic extraction: identical keyword-matching logic used to build the
    benchmark thresholds. Removes LLM judgment from the counting step entirely.
    Bounded to the first TRUNC_LEN characters, matching skill_gap_agent_analysis.ipynb."""
    tech, prog, soft, _tools = extract_skills_from_text(cv_text[:TRUNC_LEN])
    # Severity tracks tech readiness — soft skills excluded so they can't mask a real tech gap.
    total = len(tech) + len(prog)
    t = category_thresholds.get(job_category, category_thresholds["__global__"])
    # Fewer skills than most postings in this category require -> bigger gap -> High severity.
    if total <= t["q33"]:
        severity = "High"
    elif total <= t["q66"]:
        severity = "Medium"
    else:
        severity = "Low"
    return {
        "extracted_technical_skills": tech,
        "extracted_programming_languages": prog,
        "extracted_soft_skills": soft,
        "total_skill_count": total,
        "computed_severity": severity,
    }


def _dispatch_tool_call(tool_call, category_thresholds, cv_text):
    name = tool_call.function.name
    try:
        args = json.loads(tool_call.function.arguments or "{}")
    except json.JSONDecodeError:
        args = {}

    if name == "get_category_thresholds":
        return get_category_thresholds_tool(category_thresholds, args.get("job_category", "")), None
    elif name == "extract_skills_exact":
        return extract_skills_exact_tool(category_thresholds, args.get("job_category", ""), cv_text), None
    elif name == "submit_classification":
        return {"status": "recorded"}, args
    return {"error": f"unknown tool {name}"}, None


def run_agent(client, category_thresholds, job_category, cv_text, max_turns=6, max_retries=3, on_step=None):
    """Runs the tool-calling agent loop for one CV. Returns a result dict.
    on_step(msg: str) is called with live progress updates, if provided."""
    def step(msg):
        if on_step:
            on_step(msg)

    desc_snippet = str(cv_text)[:4000]
    user_msg = f"Target job category: {job_category}\n\nCandidate CV:\n{desc_snippet}"
    exact_result = None
    messages = [
        {"role": "system", "content": build_system_prompt()},
        {"role": "user", "content": user_msg},
    ]

    step("🧠 Sending your CV to the DeepSeek agent...")
    t0 = time.time()
    for turn in range(max_turns):
        resp = None
        for attempt in range(max_retries):
            try:
                resp = client.chat.completions.create(
                    model=MODEL, messages=messages, tools=AGENT_TOOLS,
                    tool_choice="auto", temperature=0.1,
                )
                break
            except Exception as e:
                if attempt == max_retries - 1:
                    step(f"❌ Agent call failed: {e}")
                    return {"severity": None, "error": str(e), "latency_s": time.time() - t0}
                step(f"⚠️ Retrying after a network hiccup (attempt {attempt + 2}/{max_retries})...")
                time.sleep((2 ** attempt) + random.random())

        msg = resp.choices[0].message
        messages.append(msg.model_dump(exclude_none=True))

        if not msg.tool_calls:
            step("❌ Agent did not return a structured answer.")
            return {"severity": None, "error": "no tool call returned", "latency_s": time.time() - t0}

        for tc in msg.tool_calls:
            if tc.function.name == "get_category_thresholds":
                step(f"🔧 Checking benchmark thresholds for *{job_category}*...")
            elif tc.function.name == "extract_skills_exact":
                step("🔍 Running deterministic skill extraction on your CV...")
            elif tc.function.name == "submit_classification":
                step("✅ Finalising severity classification...")

            result, final_args = _dispatch_tool_call(tc, category_thresholds, cv_text)
            messages.append({"role": "tool", "tool_call_id": tc.id, "content": json.dumps(result)})
            if tc.function.name == "extract_skills_exact":
                exact_result = result
            if tc.function.name == "submit_classification" and final_args is not None:
                step(f"🎯 Done — severity: **{final_args.get('severity')}**")
                # extracted_technical_skills is anchored to the deterministic tool's own
                # output rather than trusted from the model's free-form function-call args:
                # the model is instructed to copy extract_skills_exact verbatim but isn't
                # code-enforced to, so a tech skill (e.g. "statistical analysis") could
                # otherwise appear/disappear between runs on the same CV text.
                exact_tech = exact_result["extracted_technical_skills"] if exact_result else final_args.get("extracted_technical_skills", [])
                return {
                    "severity": final_args.get("severity"),
                    "confidence": final_args.get("confidence"),
                    "total_skill_count": final_args.get("total_skill_count"),
                    "extracted_technical_skills": exact_tech,
                    "extracted_programming_languages": final_args.get("extracted_programming_languages", []),
                    "extracted_soft_skills": final_args.get("extracted_soft_skills", []),
                    "reasoning": final_args.get("reasoning"),
                    "turns": turn + 1,
                    "latency_s": time.time() - t0,
                    "error": None,
                }

    step("❌ Agent exceeded the maximum number of reasoning turns.")
    return {"severity": None, "error": "max_turns exceeded", "latency_s": time.time() - t0}


# ---------------------------------------------------------------------------
# CV parsing (used for tools/experience/projects — descriptive only, not fed
# into the severity decision, keeping that decision purely agent-driven)
# ---------------------------------------------------------------------------
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


def parse_cv_metadata(text):
    """Experience/projects/tools detection — descriptive UI fields only."""
    tl = text.lower()
    tools = [tk for tk in TOOLS_KW if tk in tl]

    yrs = re.findall(r"(\d+)\+?\s*year", tl)
    exp = sum(int(y) for y in yrs) if yrs else 0

    senior_kw = ["senior", "lead", "principal", "architect", "director", "head of"]
    junior_kw = ["intern", "trainee", "fresher", "junior", "entry level"]
    if any(k in tl for k in senior_kw) or exp >= 5:   rating = 5
    elif exp >= 3:                                     rating = 4
    elif exp >= 1:                                      rating = 3
    elif any(k in tl for k in junior_kw):              rating = 2
    else:                                              rating = 3

    has_projects = bool(re.search(
        r"(project|built|developed|implemented|designed|deployed|created)", tl))

    return {"tools": tools, "tech_rating": rating, "exp_years": exp, "has_projects": has_projects}


def rule_based_severity(cv_text, need_all):
    """Fallback if the agent is unavailable — same keyword-coverage rule as before.
    need_all/coverage cover tech + programming only; soft is still extracted and
    returned for display but excluded from the severity/coverage math."""
    tech, prog, soft, _ = extract_skills_from_text(cv_text)
    have_all = set(tech) | set(prog)
    total_req = len(need_all)
    total_have = len(have_all & need_all)
    coverage = (total_have / total_req * 100) if total_req > 0 else 100.0
    if coverage >= 75:   severity = "Low"
    elif coverage >= 45: severity = "Medium"
    else:                severity = "High"
    return severity, tech, prog, soft, coverage


# ---------------------------------------------------------------------------
# Plotly dashboard charts
# ---------------------------------------------------------------------------
SEV_COLOR = {"Low": COLOR_GOOD, "Medium": COLOR_WARNING, "High": COLOR_CRITICAL}
SEV_INDEX = {"Low": 0, "Medium": 1, "High": 2}


def gauge_severity(severity):
    idx = SEV_INDEX[severity]
    fig = go.Figure(go.Indicator(
        mode="gauge",
        value=idx + 0.5,
        gauge={
            "shape": "angular",
            "axis": {"range": [0, 3], "tickvals": [0.5, 1.5, 2.5],
                      "ticktext": ["LOW", "MEDIUM", "HIGH"], "tickfont": {"size": 12}},
            "bar": {"color": "rgba(11,11,11,0.75)", "thickness": 0.28},
            "bgcolor": "rgba(0,0,0,0)",
            "steps": [
                {"range": [0, 1], "color": "rgba(12,163,12,.28)"},
                {"range": [1, 2], "color": "rgba(250,178,25,.30)"},
                {"range": [2, 3], "color": "rgba(208,59,59,.28)"},
            ],
            "threshold": {"line": {"color": SEV_COLOR[severity], "width": 4},
                          "thickness": 0.9, "value": idx + 0.5},
        },
        title={"text": f"<b>{severity.upper()}</b> SEVERITY", "font": {"size": 16}},
    ))
    fig.update_layout(height=230, margin=dict(l=20, r=20, t=50, b=10),
                       paper_bgcolor="rgba(0,0,0,0)", font=dict(color="#888"))
    return fig


def donut_coverage(coverage):
    fig = go.Figure(go.Pie(
        values=[coverage, max(0, 100 - coverage)],
        labels=["Have", "Gap"],
        hole=0.65,
        marker=dict(colors=[COLOR_GOOD, "rgba(195,194,183,.35)"]),
        textinfo="label+percent",
        textfont=dict(size=12),
        hovertemplate="%{label}: %{percent}<extra></extra>",
        sort=False,
    ))
    fig.update_layout(
        height=230, margin=dict(l=10, r=10, t=30, b=10),
        showlegend=False, paper_bgcolor="rgba(0,0,0,0)",
        annotations=[dict(text=f"<b>{coverage:.0f}%</b>", x=0.5, y=0.5,
                           font_size=22, showarrow=False, font_color="#888")],
    )
    return fig


def bar_skills_by_category(have_t, have_p, have_s, miss_t, miss_p, miss_s):
    cats = ["Technical", "Languages", "Soft Skills"]
    have = [len(have_t), len(have_p), len(have_s)]
    miss = [len(miss_t), len(miss_p), len(miss_s)]

    fig = go.Figure()
    fig.add_trace(go.Bar(name="Have", x=cats, y=have, marker_color=COLOR_GOOD,
                          hovertemplate="Have %{y} %{x} skill(s)<extra></extra>"))
    fig.add_trace(go.Bar(name="Missing", x=cats, y=miss, marker_color=COLOR_CRITICAL,
                          hovertemplate="Missing %{y} %{x} skill(s)<extra></extra>"))
    fig.update_layout(
        barmode="group", bargap=0.3, bargroupgap=0.15,
        height=320, margin=dict(l=10, r=10, t=30, b=10),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        yaxis=dict(gridcolor="rgba(150,150,150,.15)", zeroline=False),
        font=dict(color="#888"),
    )
    return fig


# ---------------------------------------------------------------------------
# Main app
# ---------------------------------------------------------------------------
def main():
    career_profiles, category_thresholds = load_benchmarks()
    career_options = sorted(career_profiles.keys())
    total_postings = sum(p["total_postings"] for p in career_profiles.values())

    st.markdown(
        f"""
        <div class="hero-wrap">
            <span class="hero-badge">🤖 Agentic Career Intelligence</span>
            <div class="main-title">Skill Gap Analyser — Agent Edition</div>
            <div class="subtitle">Select your target career, upload your CV — a DeepSeek reasoning
            agent (via TensorX) reads it live and explains your skill gap in plain language.</div>
            <div class="stat-row">
                <div class="stat-chip"><span class="num">84.2%</span><span class="lbl">Agent Accuracy</span></div>
                <div class="stat-chip"><span class="num">{len(career_options)}</span><span class="lbl">Career Paths</span></div>
                <div class="stat-chip"><span class="num">{total_postings:,}</span><span class="lbl">Postings Analysed</span></div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.write("")

    st.markdown(
        """
        <div class="steps-row">
            <div class="step-card2"><span class="step-num2">1</span><span class="step-icon">🎯</span>
                <h4>Pick a Target Career</h4><p>Choose from 8 in-demand tech roles, benchmarked against real job postings.</p></div>
            <div class="step-card2"><span class="step-num2">2</span><span class="step-icon">📄</span>
                <h4>Upload Your CV</h4><p>PDF, DOCX, or plain text — or just paste it in directly.</p></div>
            <div class="step-card2"><span class="step-num2">3</span><span class="step-icon">🧠</span>
                <h4>Get Your Analysis</h4><p>A reasoning agent scores your skill gap and builds a learning plan.</p></div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.divider()

    with st.sidebar:
        st.header("⚙️ Settings")

        api_key = os.environ.get("TENSORX_API_KEY")
        if not api_key:
            try:
                api_key = st.secrets.get("TENSORX_API_KEY")
            except Exception:
                api_key = None
        if not api_key:
            api_key = st.text_input("TensorX API key", type="password",
                                     help="Not stored — only kept for this session.")
        agent_available = bool(api_key)

        dot_cls = "pulse-dot" if agent_available else "pulse-dot off"
        status_txt = "Agent online — DeepSeek via TensorX" if agent_available else "Agent offline — rule-based fallback"
        st.markdown(f'<span class="{dot_cls}"></span>{status_txt}', unsafe_allow_html=True)

        st.markdown("---")
        selected_career = st.selectbox(
            "🎯 Target Career", career_options, index=0,
            help="Select the career you are targeting or hiring for.",
        )

        profile = career_profiles[selected_career]
        st.markdown("---")
        st.markdown(f"**📋 Benchmark — {selected_career}**")

        sb1, sb2 = st.columns(2)
        sb1.markdown(
            f'<div class="metric-card"><div class="metric-val" style="font-size:1.25rem">{profile["total_postings"]}</div>'
            f'<div class="metric-lbl">Postings</div></div>', unsafe_allow_html=True)
        sal_txt = f'${profile["avg_salary"]:,}' if profile["avg_salary"] else "—"
        sb2.markdown(
            f'<div class="metric-card"><div class="metric-val" style="font-size:1.25rem">{sal_txt}</div>'
            f'<div class="metric-lbl">Avg Salary</div></div>', unsafe_allow_html=True)

        st.write("")
        if profile["required_tech"]:
            st.markdown("**Top Required Skills**")
            st.markdown(" ".join(f'<span class="skill-have">{s.title()}</span>' for s in profile["required_tech"][:6]),
                        unsafe_allow_html=True)

        if profile["required_prog"]:
            st.markdown("**Top Languages**")
            st.markdown(" ".join(f'<span class="skill-bonus">{p.title()}</span>' for p in profile["required_prog"][:5]),
                        unsafe_allow_html=True)

        st.markdown("---")
        st.caption("📎 Supported formats: PDF · DOCX · TXT")

    st.markdown("### 📄 Upload Your CV")
    st.caption("Your CV is analysed and sent to the DeepSeek agent (via TensorX) — not stored anywhere.")

    tab_file, tab_paste = st.tabs(["📁 Upload File", "📋 Paste Text"])
    with tab_file:
        cv_file = st.file_uploader("Upload CV (PDF / DOCX / TXT)", type=["pdf", "docx", "doc", "txt"],
                                    label_visibility="collapsed")
        st.caption("Drag and drop a file, or click to browse.")
    with tab_paste:
        cv_paste = st.text_area("Or paste your CV text", height=200,
                                 placeholder="Paste your resume / CV content here...",
                                 label_visibility="collapsed")

    st.divider()
    run = st.button("🔍 Analyse My Skill Gap", type="primary", width="stretch")

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

        profile = career_profiles[selected_career]
        job_category = CAREER_TO_JOB_CATEGORY.get(selected_career, "Other Tech")
        meta = parse_cv_metadata(cv_text)

        need_tech = set(profile["required_tech"])
        need_prog = set(profile["required_prog"])
        need_soft = set(profile["required_soft"])
        # Soft skills excluded from need_all: severity/coverage/bonus track tech
        # readiness only. need_soft/have_soft are still compared separately below
        # purely for the informational Skills Breakdown display.
        need_all  = need_tech | need_prog

        agent_result = None
        if agent_available:
            with st.status("🤖 Agent analysing your CV...", expanded=True) as status_box:
                def on_step(msg):
                    status_box.write(msg)
                client = get_client(api_key)
                agent_result = run_agent(client, category_thresholds, job_category, cv_text, on_step=on_step)
                if agent_result and agent_result.get("severity"):
                    status_box.update(label="✅ Analysis complete", state="complete", expanded=False)
                else:
                    status_box.update(label="⚠️ Agent unavailable — falling back", state="error", expanded=False)

        if agent_result and agent_result.get("severity"):
            severity   = agent_result["severity"]
            have_tech  = set(s.lower() for s in agent_result["extracted_technical_skills"])
            have_prog  = set(s.lower() for s in agent_result["extracted_programming_languages"])
            have_soft  = set(s.lower() for s in agent_result["extracted_soft_skills"])
            using_agent = True
        else:
            if agent_available and agent_result:
                st.warning(f"⚠️ Agent call failed ({agent_result.get('error')}) — using rule-based fallback.")
            severity, tech, prog, soft, _ = rule_based_severity(cv_text, need_all)
            have_tech, have_prog, have_soft = set(tech), set(prog), set(soft)
            using_agent = False

        miss_tech = sorted(need_tech - have_tech)
        miss_prog = sorted(need_prog - have_prog)
        miss_soft = sorted(need_soft - have_soft)
        matched_tech = sorted(have_tech & need_tech)
        matched_prog = sorted(have_prog & need_prog)
        matched_soft = sorted(have_soft & need_soft)

        # have_all/coverage/all_missing/bonus stay tech+prog only, matching need_all
        # above — soft skills (matched_soft/miss_soft) remain available for the
        # Skills Breakdown tab but don't factor into severity or the headline metrics.
        have_all   = have_tech | have_prog
        total_req  = len(need_all)
        total_have = len(have_all & need_all)
        coverage   = (total_have / total_req * 100) if total_req > 0 else 100.0
        all_missing = miss_tech + miss_prog

        st.divider()
        st.header(f"📊 Results — {selected_career}")
        st.caption(f"Analysed at {datetime.now().strftime('%H:%M:%S')}")

        sev_msg = {
            "Low":    "✅ Strong match — minor gaps only. You're nearly job-ready!",
            "Medium": "⚠️ Moderate gap — targeted upskilling will make you competitive.",
            "High":   "❌ Significant gap — a focused learning plan is recommended.",
        }
        st.markdown(
            f'<div class="sev-banner sev-{severity}"><strong>Skill Gap Severity: {severity.upper()}</strong><br>'
            f'{sev_msg[severity]}</div>',
            unsafe_allow_html=True,
        )
        if using_agent:
            conf = agent_result.get("confidence")
            conf_str = f" · confidence {conf:.0%}" if isinstance(conf, (int, float)) else ""
            st.caption(f"🤖 Predicted by **DeepSeek Agent** via TensorX{conf_str}")
        else:
            st.caption("📐 Calculated by rule-based coverage threshold (agent unavailable)")

        tab_overview, tab_skills, tab_courses, tab_agent = st.tabs(
            ["📊 Overview", "🧩 Skills Breakdown", "📚 Recommendations", "🧠 Agent Reasoning"]
        )

        # --- Overview tab ---
        with tab_overview:
            m1, m2, m3, m4 = st.columns(4)
            for col, val, lbl in [
                (m1, f"{coverage:.0f}%",          "Skill Coverage"),
                (m2, f"{total_have}/{total_req}",  "Skills Matched"),
                (m3, f"{len(all_missing)}",        "Skills to Acquire"),
                (m4, f"{meta['tech_rating']}/5",   "Est. Seniority"),
            ]:
                col.markdown(
                    f'<div class="metric-card"><div class="metric-val">{val}</div>'
                    f'<div class="metric-lbl">{lbl}</div></div>',
                    unsafe_allow_html=True,
                )

            st.write("")
            gc1, gc2 = st.columns(2)
            with gc1:
                st.plotly_chart(gauge_severity(severity), config={"displayModeBar": False})
            with gc2:
                st.plotly_chart(donut_coverage(coverage), config={"displayModeBar": False})

            if not meta["has_projects"]:
                st.warning("⚠️ No projects detected in your CV — adding a personal project will significantly strengthen your profile.")
            if meta["exp_years"] > 0:
                st.info(f"📅 Experience detected: approximately **{meta['exp_years']} year(s)**")

        # --- Skills Breakdown tab ---
        with tab_skills:
            st.plotly_chart(
                bar_skills_by_category(matched_tech, matched_prog, matched_soft, miss_tech, miss_prog, miss_soft),
                config={"displayModeBar": False},
            )

            col_have, col_miss = st.columns(2, gap="large")
            with col_have:
                st.subheader("✅ Skills You Already Have")
                if matched_tech:
                    st.markdown("**Technical Skills**")
                    st.markdown(" ".join(f'<span class="skill-have">{s.title()}</span>' for s in matched_tech),
                                unsafe_allow_html=True)
                if matched_prog:
                    st.markdown("**Programming Languages**")
                    st.markdown(" ".join(f'<span class="skill-have">{p.title()}</span>' for p in matched_prog),
                                unsafe_allow_html=True)
                if matched_soft:
                    st.markdown("**Soft Skills**")
                    st.markdown(" ".join(f'<span class="skill-have">{s.title()}</span>' for s in matched_soft),
                                unsafe_allow_html=True)
                extra = sorted(have_all - need_all)  # tech + prog only, matched_soft is shown separately above
                if extra:
                    st.markdown("**Bonus Skills (not required but useful)**")
                    st.markdown(" ".join(f'<span class="skill-bonus">{s.title()}</span>' for s in extra),
                                unsafe_allow_html=True)
                if meta["tools"]:
                    st.markdown("**Tools Detected**")
                    st.markdown(" ".join(f'<span class="skill-have">{t}</span>' for t in meta["tools"][:12]),
                                unsafe_allow_html=True)

            with col_miss:
                st.subheader("❌ Skills You Still Need")
                if not all_missing:
                    st.success("🎉 You meet all the benchmark requirements!")
                else:
                    if miss_tech:
                        st.markdown("**Technical Skills**")
                        st.markdown(" ".join(f'<span class="skill-missing">{s.title()}</span>' for s in miss_tech),
                                    unsafe_allow_html=True)
                    if miss_prog:
                        st.markdown("**Programming Languages**")
                        st.markdown(" ".join(f'<span class="skill-missing">{p.title()}</span>' for p in miss_prog),
                                    unsafe_allow_html=True)
                    if miss_soft:
                        st.markdown("**Soft Skills**")
                        st.markdown(" ".join(f'<span class="skill-missing">{s.title()}</span>' for s in miss_soft),
                                    unsafe_allow_html=True)

            if all_missing:
                st.write("")
                st.subheader("🎯 Priority Skills to Acquire")
                for i, skill in enumerate(all_missing[:8], 1):
                    st.markdown(
                        f'<div class="priority-chip"><span class="priority-num">#{i}</span>{skill.title()}</div>',
                        unsafe_allow_html=True,
                    )

        # --- Recommendations tab ---
        with tab_courses:
            if all_missing:
                timeline = {"High": "10–14 weeks intensive", "Medium": "4–8 weeks targeted", "Low": "2–4 weeks polish"}
                st.info(f"⏱️ Suggested learning timeline: **{timeline[severity]}**")
                st.subheader("📚 Course Recommendations")
                for skill in all_missing[:10]:
                    course = COURSES.get(skill.lower())
                    title, platform, free = course if course else (f"Search: {skill.title()} course", "Coursera / Udemy / YouTube", True)
                    badge = '<span class="badge-free">FREE</span>' if free else '<span class="badge-paid">PAID</span>'
                    st.markdown(
                        f'<div class="course-card"><strong>🔧 {skill.title()}</strong><br>'
                        f'📖 {title}&nbsp;&nbsp;{badge}<span style="color:#888;font-size:.82rem">&nbsp;— {platform}</span></div>',
                        unsafe_allow_html=True,
                    )
            else:
                st.success("🎉 No missing skills — no courses needed right now!")

            st.write("")
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
            for i, step_text in enumerate(plans[severity], 1):
                st.markdown(f'<div class="step-box"><strong>{i}.</strong> {step_text}</div>', unsafe_allow_html=True)

            st.write("")
            with st.expander("📄 Full Text Summary"):
                lines = [
                    "=" * 60, "     SKILL GAP ANALYSIS REPORT (Agent Edition)", "=" * 60,
                    f"  Target Career  : {selected_career}",
                    f"  Benchmark      : {profile['total_postings']} real job postings",
                    f"  Severity Engine: {'DeepSeek Agent (TensorX)' if using_agent else 'Rule-based fallback'}",
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
                if using_agent and agent_result.get("reasoning"):
                    lines.insert(8, "  Agent reasoning: " + agent_result["reasoning"])
                st.code("\n".join(lines), language=None)

        # --- Agent Reasoning tab ---
        with tab_agent:
            if using_agent:
                st.markdown(
                    f'<div class="reasoning-box">🧠 <strong>Agent reasoning:</strong> {agent_result.get("reasoning")}</div>',
                    unsafe_allow_html=True,
                )
                rc1, rc2, rc3 = st.columns(3)
                conf = agent_result.get("confidence")
                rc1.markdown(f'<div class="metric-card"><div class="metric-val">{conf:.0%}</div>'
                             f'<div class="metric-lbl">Confidence</div></div>' if isinstance(conf, (int, float)) else "",
                             unsafe_allow_html=True)
                rc2.markdown(f'<div class="metric-card"><div class="metric-val">{agent_result.get("turns", "—")}</div>'
                             f'<div class="metric-lbl">Reasoning Turns</div></div>', unsafe_allow_html=True)
                lat = agent_result.get("latency_s")
                rc3.markdown(f'<div class="metric-card"><div class="metric-val">{lat:.1f}s</div>'
                             f'<div class="metric-lbl">Response Time</div></div>' if lat else "",
                             unsafe_allow_html=True)

                st.write("")
                st.markdown("**Raw extracted skills (agent's own semantic reading):**")
                st.json({
                    "technical": agent_result.get("extracted_technical_skills", []),
                    "programming_languages": agent_result.get("extracted_programming_languages", []),
                    "soft_skills": agent_result.get("extracted_soft_skills", []),
                    "total_skill_count": agent_result.get("total_skill_count"),
                })

                st.write("")
                st.caption(
                    "ℹ️ **Evaluation note:** on the full 663-row held-out test set (see "
                    "`skill_gap_agent_analysis.ipynb`), this tool-augmented agent — which delegates the "
                    "skill count to the same deterministic keyword extraction used to build the benchmark "
                    "thresholds, rather than judging skill presence itself — reached **84.2% accuracy** "
                    "(87.6% precision, 80.5% recall), ahead of Random Forest (72.4%), SVM (67.0%), and "
                    "BERT (52.0%) on the same task. Its main remaining error source is the extraction tool's "
                    "bounded context window, which can miss skills mentioned very late in long postings."
                )
            else:
                st.info(
                    "The agent was not used for this analysis (no TensorX API key, or the agent call "
                    "failed) — severity was calculated with a simple rule-based keyword-coverage "
                    "threshold instead. Add a TensorX API key in the sidebar to see live agent reasoning."
                )

    st.divider()
    st.markdown(
        "<div style='text-align:center;color:#888;font-size:.8rem'>"
        "Skill Gap Analyser — Agent Edition · Powered by DeepSeek via TensorX · 2026"
        "<br>© 2026 Srikaran Sankar. All rights reserved."
        "</div>",
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
