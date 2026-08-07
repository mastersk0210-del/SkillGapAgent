"""One-off script: precompute agent_benchmarks.json (career_profiles +
category_thresholds) from postings.csv, mirroring load_benchmarks() in
agent_app.py exactly, so the deployed app doesn't need the 516MB CSV.
"""
import os, re, json
from collections import Counter
import pandas as pd

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(SCRIPT_DIR, "postings.csv")
OUT_FILE = os.path.join(SCRIPT_DIR, "agent_benchmarks.json")

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

SKILL_ALIASES = {
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
    "leadership":                    ["team lead", "led a team", "managed a team"],
    "teamwork":                      ["team player", "cross functional team", "collaborative team"],
    "critical thinking":             ["critical analysis", "problem analysis"],
    "creativity":                    ["creative thinking", "innovative"],
    "collaboration":                 ["cross functional collaboration", "team collaboration"],
    "presentation":                  ["public speaking", "stakeholder presentation"],
    "project management":           ["pmp", "agile", "scrum", "project lead", "program management"],
    "analytical thinking":          ["analytical skills", "data driven decision making"],
    "attention to detail":          ["detail oriented", "meticulous"],
    "problem solving":               ["problem solving skills"],
    "time management":               ["prioritization", "deadline management"],
    "adaptability":                  ["flexibility", "adaptable"],
}
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


def main():
    df = pd.read_csv(DATA_FILE, low_memory=False)
    mask = df["title"].str.lower().apply(lambda x: any(t in str(x) for t in TECH_TITLES))
    df = df[mask].copy().reset_index(drop=True)

    df["career"] = df["title"].apply(map_career)
    df["job_category"] = df["title"].apply(map_job_category)

    extracted = df["description"].apply(extract_skills_from_text)
    df["tech_skills"] = extracted.apply(lambda x: x[0])
    df["prog_langs"]  = extracted.apply(lambda x: x[1])
    df["soft_skills"] = extracted.apply(lambda x: x[2])
    df["tools"]       = extracted.apply(lambda x: x[3])
    df["total_skills"] = df["tech_skills"].apply(len) + df["prog_langs"].apply(len)

    career_profiles = {}
    for career, grp in df[df["career"].notna()].groupby("career"):
        n = len(grp)
        threshold = max(1, int(n * 0.15))

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

    thresholds = df.groupby("job_category")["total_skills"].quantile([0.33, 0.66]).unstack()
    thresholds.columns = ["q33", "q66"]
    global_q33 = df["total_skills"].quantile(0.33)
    global_q66 = df["total_skills"].quantile(0.66)

    category_thresholds = {
        cat: {"q33": float(row["q33"]), "q66": float(row["q66"])}
        for cat, row in thresholds.iterrows()
    }
    category_thresholds["__global__"] = {"q33": float(global_q33), "q66": float(global_q66)}

    out = {"career_profiles": career_profiles, "category_thresholds": category_thresholds}
    with open(OUT_FILE, "w") as f:
        json.dump(out, f, indent=2)
    print(f"Wrote {OUT_FILE} ({os.path.getsize(OUT_FILE)} bytes)")
    for k, v in career_profiles.items():
        print(f"  {k}: {v['total_postings']} postings")
    print("category_thresholds:", list(category_thresholds.keys()))


if __name__ == "__main__":
    main()
