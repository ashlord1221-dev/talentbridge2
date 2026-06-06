import os, sqlite3, json, re, math
from datetime import datetime
from flask import (Flask, render_template, request, redirect, url_for,
                   session, flash, jsonify, g)
from werkzeug.utils import secure_filename
from werkzeug.security import generate_password_hash, check_password_hash

# ── Vercel: force writable /tmp paths ─────────────────────────────────────
if os.environ.get("VERCEL") or not os.access(os.path.dirname(os.path.abspath(__file__)), os.W_OK):
    os.environ.setdefault("DB_PATH", "/tmp/talentbridge.db")
    os.environ.setdefault("UPLOAD_FOLDER", "/tmp/uploads")
os.makedirs(os.environ.get("UPLOAD_FOLDER", "/tmp/uploads"), exist_ok=True)

try:
    from pdfminer.high_level import extract_text as pdf_extract
    PDF_OK = True
except ImportError:
    PDF_OK = False

try:
    import docx as docx_lib
    DOCX_OK = True
except ImportError:
    DOCX_OK = False

# ── App setup ──────────────────────────────────────────────────────────────
BASE = os.path.dirname(os.path.abspath(__file__))
app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "talentbridge-dev-key-2025")

UPLOAD_FOLDER = os.environ.get("UPLOAD_FOLDER", os.path.join(BASE, "static", "uploads"))
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["TEMPLATES_AUTO_RELOAD"] = True
app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024
ALLOWED = {"pdf", "doc", "docx", "txt"}

DB_PATH = os.environ.get("DB_PATH", os.path.join(BASE, "talentbridge.db"))

# ── Database ───────────────────────────────────────────────────────────────
def get_db():
    if "db" not in g:
        g.db = sqlite3.connect(DB_PATH)
        g.db.row_factory = sqlite3.Row
        g.db.execute("PRAGMA journal_mode=WAL")
        g.db.execute("PRAGMA foreign_keys=ON")
    return g.db

@app.teardown_appcontext
def close_db(e=None):
    db = g.pop("db", None)
    if db:
        db.close()

def query(sql, params=(), one=False, commit=False):
    db = get_db()
    cur = db.execute(sql, params)
    if commit:
        db.commit()
        return cur.lastrowid
    return (cur.fetchone() if one else cur.fetchall())

# ── Schema ─────────────────────────────────────────────────────────────────
SCHEMA = """
CREATE TABLE IF NOT EXISTS candidates (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL, email TEXT UNIQUE NOT NULL, password TEXT NOT NULL,
    phone TEXT, location TEXT, skills TEXT, experience INTEGER DEFAULT 0,
    education TEXT, bio TEXT, resume_file TEXT, resume_text TEXT,
    linkedin TEXT, expected_salary TEXT, work_pref TEXT DEFAULT 'Any',
    created_at TEXT DEFAULT (datetime('now'))
);
CREATE TABLE IF NOT EXISTS companies (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL, email TEXT UNIQUE NOT NULL, password TEXT NOT NULL,
    industry TEXT, size TEXT, location TEXT, website TEXT, about TEXT,
    logo_letter TEXT, logo_color TEXT DEFAULT '#2563EB', logo_bg TEXT DEFAULT '#dbeafe',
    created_at TEXT DEFAULT (datetime('now'))
);
CREATE TABLE IF NOT EXISTS jobs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    company_id INTEGER NOT NULL, title TEXT NOT NULL, department TEXT,
    job_type TEXT DEFAULT 'Full-time', work_mode TEXT DEFAULT 'On-site',
    location TEXT, salary TEXT, experience_req INTEGER DEFAULT 0,
    skills_req TEXT, description TEXT, requirements TEXT, deadline TEXT,
    status TEXT DEFAULT 'Active', featured INTEGER DEFAULT 0,
    created_at TEXT DEFAULT (datetime('now')),
    FOREIGN KEY(company_id) REFERENCES companies(id)
);
CREATE TABLE IF NOT EXISTS applications (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    job_id INTEGER NOT NULL, candidate_id INTEGER NOT NULL,
    cover_letter TEXT, expected_salary TEXT, notice_period TEXT,
    stage TEXT DEFAULT 'Applied', ai_score REAL DEFAULT 0,
    score_breakdown TEXT, applied_at TEXT DEFAULT (datetime('now')),
    UNIQUE(job_id, candidate_id),
    FOREIGN KEY(job_id) REFERENCES jobs(id),
    FOREIGN KEY(candidate_id) REFERENCES candidates(id)
);
CREATE TABLE IF NOT EXISTS saved_jobs (
    candidate_id INTEGER, job_id INTEGER,
    PRIMARY KEY(candidate_id, job_id)
);
"""

def init_db():
    with sqlite3.connect(DB_PATH) as db:
        for stmt in SCHEMA.strip().split(";"):
            s = stmt.strip()
            if s:
                db.execute(s)
        db.commit()
        seed_demo_data(db)

def seed_demo_data(db):
    comps = [
        ("Google India","google@demo.com","demo123","Technology","50,000+","Bangalore","https://google.com","Building for everyone.","G","#4285f4","#dbeafe"),
        ("Zomato","zomato@demo.com","demo123","Food Tech","5,000+","Gurugram","https://zomato.com","Delivering happiness.","Z","#dc2626","#fee2e2"),
        ("Infosys","infosys@demo.com","demo123","IT Services","300,000+","Bangalore","https://infosys.com","Navigate your next.","I","#16a34a","#dcfce7"),
        ("Flipkart","flipkart@demo.com","demo123","E-commerce","30,000+","Bangalore","https://flipkart.com","India's marketplace.","F","#ca8a04","#fef9c3"),
        ("BYJU'S","byjus@demo.com","demo123","EdTech","10,000+","Bangalore","https://byjus.com","Think & Learn.","B","#7c3aed","#ede9fe"),
    ]
    for c in comps:
        try:
            db.execute("INSERT OR IGNORE INTO companies(name,email,password,industry,size,location,website,about,logo_letter,logo_color,logo_bg) VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                (c[0],c[1],generate_password_hash(c[2]),c[3],c[4],c[5],c[6],c[7],c[8],c[9],c[10]))
        except: pass
    db.commit()

    jobs = [
        (1,"Senior Frontend Engineer","Engineering","Full-time","Remote","Bangalore","₹40–60 LPA",5,"React,TypeScript,Next.js,GraphQL,CSS","Build world-class UIs for Google's consumer products using cutting-edge stacks.","5+ years React\nTypeScript proficiency\nGraphQL experience\nStrong CSS skills","2025-12-31",1),
        (1,"ML Engineer","Data","Full-time","Remote","Hyderabad","₹50–80 LPA",5,"Python,TensorFlow,PyTorch,MLOps,Spark","Build and deploy ML models powering Google's most critical features.","5+ years ML\nPython expert\nTF or PyTorch\nMLOps experience","2025-12-31",1),
        (2,"Data Scientist","Data","Full-time","Hybrid","Mumbai","₹25–40 LPA",3,"Python,ML,SQL,TensorFlow,Pandas","Drive data-driven insights for Zomato's growth.","3+ years experience\nPython and SQL expert\nML modelling\nStatistics background","2025-11-30",0),
        (2,"Marketing Manager","Marketing","Full-time","Hybrid","Delhi","₹15–25 LPA",3,"Digital Marketing,Analytics,Content,SEO","Lead growth marketing for Zomato.","3+ years digital marketing\nAnalytics tools\nCampaign management","2025-11-30",0),
        (3,"Product Manager","Product","Full-time","On-site","Chennai","₹30–45 LPA",4,"Agile,Roadmapping,Analytics,Jira,SQL","Lead product strategy for enterprise software at Infosys.","4+ years PM\nAgile certified\nSQL familiarity\nStakeholder management","2025-12-15",0),
        (3,"Java Backend Developer","Engineering","Full-time","On-site","Bangalore","₹20–35 LPA",3,"Java,Spring Boot,Microservices,Docker,PostgreSQL","Build scalable microservices for Infosys enterprise clients.","3+ years Java\nSpring Boot\nREST APIs\nDocker","2025-12-15",0),
        (4,"Full Stack Developer","Engineering","Full-time","Hybrid","Bangalore","₹20–35 LPA",2,"React,Node.js,MongoDB,Docker,REST APIs","Build and scale India's largest e-commerce infrastructure.","2+ years full-stack\nReact and Node.js\nMongoDB or PostgreSQL\nDocker","2025-11-15",1),
        (4,"DevOps Engineer","Engineering","Full-time","On-site","Bangalore","₹22–38 LPA",4,"Kubernetes,AWS,Terraform,CI/CD,Linux","Scale infrastructure for millions of daily transactions.","4+ years DevOps\nAWS or GCP\nKubernetes\nTerraform","2025-12-01",0),
        (5,"UX Designer","Design","Full-time","Remote","Bangalore","₹18–28 LPA",3,"Figma,User Research,Prototyping,Motion Design","Design learning experiences for 150M+ students at BYJU'S.","3+ years UX\nFigma expert\nUser research\nA/B testing","2025-11-20",0),
        (5,"Content Writer","Marketing","Full-time","Remote","Anywhere","₹8–14 LPA",1,"Content Writing,SEO,Research,Editing","Create educational content for BYJU'S learning platform.","1+ year content writing\nSEO knowledge\nExcellent English","2025-11-10",0),
    ]
    for j in jobs:
        try:
            db.execute("INSERT OR IGNORE INTO jobs(company_id,title,department,job_type,work_mode,location,salary,experience_req,skills_req,description,requirements,deadline,featured) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)", j)
        except: pass
    db.commit()

    cands = [
        ("Arjun Sharma","arjun@demo.com","demo123","9876543210","Chennai","React,TypeScript,Node.js,GraphQL,CSS,Docker,Jest",5,"B.Tech CS","Senior frontend developer with 5 years building scalable web applications.","","","https://linkedin.com/in/arjun","25–35 LPA","Remote"),
        ("Priya Nair","priya@demo.com","demo123","9876543211","Bangalore","React,TypeScript,Python,Docker,GraphQL,Redux,AWS",7,"M.Tech CS","Full-stack engineer with expertise in cloud-native applications.","","","","35–50 LPA","Hybrid"),
    ]
    for c in cands:
        try:
            db.execute("INSERT OR IGNORE INTO candidates(name,email,password,phone,location,skills,experience,education,bio,resume_file,resume_text,linkedin,expected_salary,work_pref) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                (c[0],c[1],generate_password_hash(c[2]),c[3],c[4],c[5],c[6],c[7],c[8],c[9],c[10],c[11],c[12],c[13]))
        except: pass
    db.commit()

# ── Helpers ────────────────────────────────────────────────────────────────
def allowed_file(fname):
    return "." in fname and fname.rsplit(".",1)[1].lower() in ALLOWED

def extract_resume_text(filepath):
    ext = filepath.rsplit(".",1)[-1].lower()
    try:
        if ext == "pdf" and PDF_OK:
            return pdf_extract(filepath) or ""
        elif ext in ("doc","docx") and DOCX_OK:
            doc = docx_lib.Document(filepath)
            return "\n".join(p.text for p in doc.paragraphs)
        elif ext == "txt":
            return open(filepath,"r",errors="ignore").read()
    except: pass
    return ""

SKILL_KEYWORDS = [
    "python","java","javascript","typescript","react","angular","vue","node","nodejs",
    "django","flask","spring","express","fastapi","graphql","rest","sql","mysql",
    "postgresql","mongodb","redis","docker","kubernetes","aws","gcp","azure","terraform",
    "git","linux","figma","css","html","tensorflow","pytorch","pandas","numpy",
    "scikit","machine learning","deep learning","nlp","spark","tableau","powerbi",
    "agile","scrum","jira","devops","kotlin","swift","flutter","react native",
    "next.js","nextjs","webpack","redux","statistics","data analysis",
]

def parse_skills_from_text(text):
    found, low = [], text.lower()
    for sk in SKILL_KEYWORDS:
        if sk in low and sk not in found:
            found.append(sk.title() if len(sk)>3 else sk.upper())
    return ", ".join(found[:25])

def compute_ai_score(candidate, job):
    c_skills = set(s.strip().lower() for s in (candidate["skills"] or "").split(",") if s.strip())
    j_skills = set(s.strip().lower() for s in (job["skills_req"] or "").split(",") if s.strip())
    if candidate["resume_text"]:
        c_skills |= set(s.strip().lower() for s in parse_skills_from_text(candidate["resume_text"]).split(",") if s.strip())
    skill_score = (len(c_skills & j_skills) / len(j_skills)) if j_skills else 0.7
    matched = c_skills & j_skills
    c_exp, j_exp = int(candidate["experience"] or 0), int(job["experience_req"] or 0)
    if j_exp == 0: exp_score = 1.0
    elif c_exp >= j_exp+3: exp_score = 1.0
    elif c_exp >= j_exp: exp_score = 0.9
    elif c_exp >= j_exp-1: exp_score = 0.65
    else: exp_score = max(0, c_exp/max(j_exp,1)*0.6)
    edu = (candidate["education"] or "").lower()
    edu_score = 1.0 if any(x in edu for x in ["phd","m.tech","mtech","m.sc","mca","ms "]) else \
                0.8 if any(x in edu for x in ["b.tech","btech","b.e","be ","b.sc","bsc","bca"]) else \
                0.6 if edu else 0.5
    pts = (40 if candidate["resume_file"] else 0) + (20 if candidate["bio"] else 0) + \
          (20 if candidate["linkedin"] else 0) + (10 if candidate["phone"] else 0) + \
          (10 if candidate["location"] else 0)
    profile_score = min(pts/100, 1.0)
    total = round(min((skill_score*0.40 + exp_score*0.25 + edu_score*0.15 + profile_score*0.20)*100, 100), 1)
    return total, {
        "skills": round(skill_score*100,1), "experience": round(exp_score*100,1),
        "education": round(edu_score*100,1), "profile": round(profile_score*100,1),
        "matched_skills": list(matched), "missing_skills": list(j_skills-c_skills),
    }

def rank_applications(job_id):
    apps = query("""SELECT a.*, c.name, c.skills, c.experience, c.education,
        c.resume_file, c.resume_text, c.bio, c.linkedin, c.phone, c.location
        FROM applications a JOIN candidates c ON a.candidate_id=c.id WHERE a.job_id=?""", (job_id,))
    job = query("SELECT * FROM jobs WHERE id=?", (job_id,), one=True)
    results = []
    for app in apps:
        score, bd = compute_ai_score(app, job)
        query("UPDATE applications SET ai_score=?, score_breakdown=? WHERE id=?",
              (score, json.dumps(bd), app["id"]), commit=True)
        results.append((app, score, bd))
    return sorted(results, key=lambda x: x[1], reverse=True)

def current_candidate():
    if session.get("user_type") == "candidate":
        return query("SELECT * FROM candidates WHERE id=?", (session["user_id"],), one=True)
    return None

def current_company():
    if session.get("user_type") == "company":
        return query("SELECT * FROM companies WHERE id=?", (session["user_id"],), one=True)
    return None

def login_required_candidate(f):
    from functools import wraps
    @wraps(f)
    def dec(*a, **kw):
        if not current_candidate():
            flash("Please log in as a candidate.","warning")
            return redirect(url_for("candidate_login"))
        return f(*a, **kw)
    return dec

def login_required_company(f):
    from functools import wraps
    @wraps(f)
    def dec(*a, **kw):
        if not current_company():
            flash("Please log in as a company.","warning")
            return redirect(url_for("company_login"))
        return f(*a, **kw)
    return dec

# ══════════════════════════════════════════════════════════════════════════
# PUBLIC ROUTES
# ══════════════════════════════════════════════════════════════════════════
@app.route("/")
def index():
    featured = query("SELECT j.*,c.name as company_name,c.logo_letter,c.logo_color,c.logo_bg FROM jobs j JOIN companies c ON j.company_id=c.id WHERE j.status='Active' AND j.featured=1 LIMIT 6")
    recent   = query("SELECT j.*,c.name as company_name,c.logo_letter,c.logo_color,c.logo_bg FROM jobs j JOIN companies c ON j.company_id=c.id WHERE j.status='Active' ORDER BY j.created_at DESC LIMIT 6")
    companies= query("SELECT c.*, COUNT(j.id) as open_jobs FROM companies c LEFT JOIN jobs j ON c.id=j.company_id AND j.status='Active' GROUP BY c.id LIMIT 6")
    stats = {
        "jobs":      query("SELECT COUNT(*) FROM jobs WHERE status='Active'", one=True)[0],
        "companies": query("SELECT COUNT(*) FROM companies", one=True)[0],
        "candidates":query("SELECT COUNT(*) FROM candidates", one=True)[0],
        "placed":    query("SELECT COUNT(*) FROM applications WHERE stage='Hired'", one=True)[0],
    }
    return render_template("index.html", featured=featured, recent=recent,
        companies=companies, stats=stats, candidate=current_candidate(), company=current_company())

@app.route("/jobs")
def jobs():
    q    = request.args.get("q","").strip()
    jtype= request.args.get("type","")
    dept = request.args.get("dept","")
    sort = request.args.get("sort","recent")
    page = int(request.args.get("page",1))
    per  = 9
    sql  = "SELECT j.*,c.name as company_name,c.logo_letter,c.logo_color,c.logo_bg FROM jobs j JOIN companies c ON j.company_id=c.id WHERE j.status='Active'"
    params = []
    if q:
        sql += " AND (j.title LIKE ? OR c.name LIKE ? OR j.skills_req LIKE ? OR j.location LIKE ?)"
        like = f"%{q}%"; params += [like,like,like,like]
    if jtype: sql += " AND j.job_type=?"; params.append(jtype)
    if dept:  sql += " AND j.department=?"; params.append(dept)
    order = {"recent":"j.created_at DESC","salary":"j.salary DESC","applicants":"(SELECT COUNT(*) FROM applications WHERE job_id=j.id) DESC"}.get(sort,"j.created_at DESC")
    sql += f" ORDER BY j.featured DESC, {order}"
    all_jobs   = query(sql, params)
    total      = len(all_jobs)
    paginated  = all_jobs[(page-1)*per:page*per]
    pages      = math.ceil(total/per)
    saved_ids  = []
    cand = current_candidate()
    if cand:
        saved_ids = [r["job_id"] for r in query("SELECT job_id FROM saved_jobs WHERE candidate_id=?", (cand["id"],))]
    return render_template("jobs.html", jobs=paginated, total=total, page=page, pages=pages,
        q=q, jtype=jtype, dept=dept, sort=sort, saved_ids=saved_ids,
        candidate=cand, company=current_company())

@app.route("/jobs/<int:job_id>")
def job_detail(job_id):
    job = query("SELECT j.*,c.name as company_name,c.logo_letter,c.logo_color,c.logo_bg,c.about,c.size,c.website FROM jobs j JOIN companies c ON j.company_id=c.id WHERE j.id=?", (job_id,), one=True)
    if not job: flash("Job not found.","danger"); return redirect(url_for("jobs"))
    app_count = query("SELECT COUNT(*) FROM applications WHERE job_id=?", (job_id,), one=True)[0]
    applied = saved = False
    cand = current_candidate()
    if cand:
        applied = bool(query("SELECT 1 FROM applications WHERE job_id=? AND candidate_id=?", (job_id,cand["id"]), one=True))
        saved   = bool(query("SELECT 1 FROM saved_jobs WHERE job_id=? AND candidate_id=?",   (job_id,cand["id"]), one=True))
    similar = query("SELECT j.*,c.name as company_name,c.logo_letter,c.logo_color,c.logo_bg FROM jobs j JOIN companies c ON j.company_id=c.id WHERE j.id!=? AND j.department=? AND j.status='Active' LIMIT 3", (job_id, job["department"]))
    return render_template("job_detail.html", job=job, app_count=app_count, applied=applied,
        saved=saved, similar=similar, candidate=cand, company=current_company())

@app.route("/companies")
def companies_page():
    comps = query("SELECT c.*, COUNT(j.id) as open_jobs FROM companies c LEFT JOIN jobs j ON c.id=j.company_id AND j.status='Active' GROUP BY c.id ORDER BY open_jobs DESC")
    return render_template("companies.html", companies=comps, candidate=current_candidate(), company=current_company())

# ══════════════════════════════════════════════════════════════════════════
# CANDIDATE AUTH
# ══════════════════════════════════════════════════════════════════════════
@app.route("/candidate/register", methods=["GET","POST"])
def candidate_register():
    if request.method == "POST":
        name=request.form.get("name","").strip(); email=request.form.get("email","").strip().lower()
        pwd=request.form.get("password",""); phone=request.form.get("phone","").strip(); loc=request.form.get("location","").strip()
        if not (name and email and pwd): flash("Name, email, and password are required.","danger"); return redirect(url_for("candidate_register"))
        if query("SELECT 1 FROM candidates WHERE email=?", (email,), one=True): flash("Email already registered.","warning"); return redirect(url_for("candidate_register"))
        uid = query("INSERT INTO candidates(name,email,password,phone,location) VALUES(?,?,?,?,?)", (name,email,generate_password_hash(pwd),phone,loc), commit=True)
        session.update({"user_id":uid,"user_type":"candidate","user_name":name})
        flash(f"Welcome, {name}! Complete your profile for better matches.","success")
        return redirect(url_for("candidate_dashboard"))
    return render_template("candidate_register.html", candidate=None, company=None)

@app.route("/candidate/login", methods=["GET","POST"])
def candidate_login():
    if request.method == "POST":
        email=request.form.get("email","").strip().lower(); pwd=request.form.get("password","")
        u = query("SELECT * FROM candidates WHERE email=?", (email,), one=True)
        if u and check_password_hash(u["password"], pwd):
            session.update({"user_id":u["id"],"user_type":"candidate","user_name":u["name"]})
            flash(f"Welcome back, {u['name']}!","success")
            return redirect(request.args.get("next") or url_for("candidate_dashboard"))
        flash("Invalid email or password.","danger")
    return render_template("candidate_login.html", candidate=None, company=None)

@app.route("/candidate/logout")
def candidate_logout():
    session.clear(); flash("Logged out successfully.","info"); return redirect(url_for("index"))

# ══════════════════════════════════════════════════════════════════════════
# CANDIDATE ACTIONS
# ══════════════════════════════════════════════════════════════════════════
@app.route("/candidate/dashboard")
@login_required_candidate
def candidate_dashboard():
    cand = current_candidate()
    apps = query("""SELECT a.*, j.title, j.salary, j.location, j.job_type, j.id as job_id,
        c.name as company_name, c.logo_letter, c.logo_color, c.logo_bg
        FROM applications a JOIN jobs j ON a.job_id=j.id JOIN companies c ON j.company_id=c.id
        WHERE a.candidate_id=? ORDER BY a.applied_at DESC""", (cand["id"],))
    saved = query("""SELECT j.*, c.name as company_name, c.logo_letter, c.logo_color, c.logo_bg
        FROM saved_jobs s JOIN jobs j ON s.job_id=j.id JOIN companies c ON j.company_id=c.id
        WHERE s.candidate_id=?""", (cand["id"],))
    recommended = query("""SELECT j.*,c.name as company_name,c.logo_letter,c.logo_color,c.logo_bg
        FROM jobs j JOIN companies c ON j.company_id=c.id
        WHERE j.status='Active' ORDER BY j.featured DESC, j.created_at DESC LIMIT 6""")
    stats = {"applied":len(apps), "shortlisted":sum(1 for a in apps if a["stage"] in ("Shortlisted","Interview","Offer","Hired")),
             "interviews":sum(1 for a in apps if a["stage"]=="Interview"), "saved":len(saved)}
    return render_template("candidate_dashboard.html", cand=cand, apps=apps, saved=saved,
        recommended=recommended, stats=stats, candidate=cand, company=None)

@app.route("/candidate/profile", methods=["GET","POST"])
@login_required_candidate
def candidate_profile():
    cand = current_candidate()
    if request.method == "POST":
        f = ["name","phone","location","skills","experience","education","bio","linkedin","expected_salary","work_pref"]
        v = {x: request.form.get(x,"").strip() for x in f}
        query("UPDATE candidates SET name=?,phone=?,location=?,skills=?,experience=?,education=?,bio=?,linkedin=?,expected_salary=?,work_pref=? WHERE id=?",
            (v["name"],v["phone"],v["location"],v["skills"],v["experience"] or 0,v["education"],v["bio"],v["linkedin"],v["expected_salary"],v["work_pref"],cand["id"]), commit=True)
        session["user_name"] = v["name"]
        flash("Profile updated!","success"); return redirect(url_for("candidate_profile"))
    return render_template("candidate_profile.html", cand=cand, candidate=cand, company=None)

@app.route("/candidate/upload-resume", methods=["GET","POST"])
@login_required_candidate
def upload_resume():
    cand = current_candidate()
    if request.method == "POST":
        file = request.files.get("resume")
        if not file or file.filename == "": flash("Please select a file.","danger"); return redirect(url_for("upload_resume"))
        if not allowed_file(file.filename): flash("Allowed: PDF, DOC, DOCX, TXT","danger"); return redirect(url_for("upload_resume"))
        fname = secure_filename(f"cand_{cand['id']}_{int(datetime.now().timestamp())}_{file.filename}")
        fpath = os.path.join(app.config["UPLOAD_FOLDER"], fname)
        file.save(fpath)
        text = extract_resume_text(fpath)
        parsed = parse_skills_from_text(text) if text else ""
        existing = set(s.strip().lower() for s in (cand["skills"] or "").split(",") if s.strip())
        new_sk   = set(s.strip().lower() for s in parsed.split(",") if s.strip())
        merged   = ", ".join(sorted(existing | new_sk))
        query("UPDATE candidates SET resume_file=?, resume_text=?, skills=? WHERE id=?",
            (fname, text, merged or cand["skills"], cand["id"]), commit=True)
        flash(f"Resume uploaded! {len(new_sk)} skills extracted.","success")
        return redirect(url_for("candidate_dashboard"))
    return render_template("upload_resume.html", cand=cand, candidate=cand, company=None)

@app.route("/jobs/<int:job_id>/apply", methods=["GET","POST"])
@login_required_candidate
def apply_job(job_id):
    cand = current_candidate()
    job  = query("SELECT j.*,c.name as company_name FROM jobs j JOIN companies c ON j.company_id=c.id WHERE j.id=?", (job_id,), one=True)
    if not job: flash("Job not found.","danger"); return redirect(url_for("jobs"))
    if query("SELECT 1 FROM applications WHERE job_id=? AND candidate_id=?", (job_id,cand["id"]), one=True):
        flash("Already applied!","info"); return redirect(url_for("job_detail", job_id=job_id))
    if request.method == "POST":
        score, bd = compute_ai_score(cand, job)
        query("INSERT INTO applications(job_id,candidate_id,cover_letter,expected_salary,notice_period,ai_score,score_breakdown) VALUES(?,?,?,?,?,?,?)",
            (job_id,cand["id"],request.form.get("cover_letter","").strip(),
             request.form.get("expected_salary","").strip(),request.form.get("notice_period","30 days"),
             score,json.dumps(bd)), commit=True)
        flash(f"Application submitted! Your AI match score is {score}%.","success")
        return redirect(url_for("candidate_dashboard"))
    score, bd = compute_ai_score(cand, job)
    return render_template("apply_job.html", job=job, cand=cand, score=score, breakdown=bd,
        candidate=cand, company=None)

@app.route("/jobs/<int:job_id>/save", methods=["POST"])
@login_required_candidate
def save_job(job_id):
    cand = current_candidate()
    exists = query("SELECT 1 FROM saved_jobs WHERE candidate_id=? AND job_id=?", (cand["id"],job_id), one=True)
    if exists:
        query("DELETE FROM saved_jobs WHERE candidate_id=? AND job_id=?", (cand["id"],job_id), commit=True)
        return jsonify({"saved":False})
    query("INSERT INTO saved_jobs(candidate_id,job_id) VALUES(?,?)", (cand["id"],job_id), commit=True)
    return jsonify({"saved":True})

# ══════════════════════════════════════════════════════════════════════════
# COMPANY AUTH
# ══════════════════════════════════════════════════════════════════════════
@app.route("/company/register", methods=["GET","POST"])
def company_register():
    if request.method == "POST":
        name=request.form.get("name","").strip(); email=request.form.get("email","").strip().lower()
        pwd=request.form.get("password",""); ind=request.form.get("industry","")
        size=request.form.get("size",""); loc=request.form.get("location","")
        if not (name and email and pwd): flash("Name, email, password required.","danger"); return redirect(url_for("company_register"))
        if query("SELECT 1 FROM companies WHERE email=?", (email,), one=True): flash("Email already registered.","warning"); return redirect(url_for("company_register"))
        import random
        lc,lb = random.choice([("#2563EB","#dbeafe"),("#dc2626","#fee2e2"),("#16a34a","#dcfce7"),("#ca8a04","#fef9c3"),("#7c3aed","#ede9fe")])
        uid = query("INSERT INTO companies(name,email,password,industry,size,location,logo_letter,logo_color,logo_bg) VALUES(?,?,?,?,?,?,?,?,?)",
            (name,email,generate_password_hash(pwd),ind,size,loc,name[0].upper(),lc,lb), commit=True)
        session.update({"user_id":uid,"user_type":"company","user_name":name})
        flash(f"Welcome, {name}! Post your first job.","success")
        return redirect(url_for("company_dashboard"))
    return render_template("company_register.html", candidate=None, company=None)

@app.route("/company/login", methods=["GET","POST"])
def company_login():
    if request.method == "POST":
        email=request.form.get("email","").strip().lower(); pwd=request.form.get("password","")
        u = query("SELECT * FROM companies WHERE email=?", (email,), one=True)
        if u and check_password_hash(u["password"], pwd):
            session.update({"user_id":u["id"],"user_type":"company","user_name":u["name"]})
            flash(f"Welcome back, {u['name']}!","success")
            return redirect(url_for("company_dashboard"))
        flash("Invalid email or password.","danger")
    return render_template("company_login.html", candidate=None, company=None)

@app.route("/company/logout")
def company_logout():
    session.clear(); flash("Logged out.","info"); return redirect(url_for("index"))

# ══════════════════════════════════════════════════════════════════════════
# COMPANY ACTIONS
# ══════════════════════════════════════════════════════════════════════════
@app.route("/company/dashboard")
@login_required_company
def company_dashboard():
    comp = current_company()
    jobs_list = query("SELECT j.*, (SELECT COUNT(*) FROM applications WHERE job_id=j.id) as app_count FROM jobs j WHERE j.company_id=? ORDER BY j.created_at DESC", (comp["id"],))
    stats = {
        "active_jobs": sum(1 for j in jobs_list if j["status"]=="Active"),
        "total_apps":  sum(j["app_count"] for j in jobs_list),
        "shortlisted": query("SELECT COUNT(*) FROM applications a JOIN jobs j ON a.job_id=j.id WHERE j.company_id=? AND a.stage IN ('Shortlisted','Interview','Offer')", (comp["id"],), one=True)[0],
        "hired":       query("SELECT COUNT(*) FROM applications a JOIN jobs j ON a.job_id=j.id WHERE j.company_id=? AND a.stage='Hired'", (comp["id"],), one=True)[0],
    }
    pipeline = {}
    for stage in ["Applied","Shortlisted","Interview","Offer","Hired","Rejected"]:
        pipeline[stage] = query("""SELECT a.*, c.name as cand_name, c.skills, c.experience, c.education, j.title as job_title
            FROM applications a JOIN candidates c ON a.candidate_id=c.id JOIN jobs j ON a.job_id=j.id
            WHERE j.company_id=? AND a.stage=? ORDER BY a.ai_score DESC LIMIT 10""", (comp["id"],stage))
    return render_template("company_dashboard.html", comp=comp, jobs=jobs_list, stats=stats,
        pipeline=pipeline, candidate=None, company=comp)

@app.route("/company/post-job", methods=["GET","POST"])
@login_required_company
def post_job():
    comp = current_company()
    if request.method == "POST":
        title = request.form.get("title","").strip()
        if not title: flash("Job title required.","danger"); return redirect(url_for("post_job"))
        f = ["title","department","job_type","work_mode","location","salary","experience_req","skills_req","description","requirements","deadline"]
        v = {x: request.form.get(x,"").strip() for x in f}
        query("INSERT INTO jobs(company_id,title,department,job_type,work_mode,location,salary,experience_req,skills_req,description,requirements,deadline,featured) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)",
            (comp["id"],v["title"],v["department"],v["job_type"],v["work_mode"],v["location"],v["salary"],
             v["experience_req"] or 0,v["skills_req"],v["description"],v["requirements"],v["deadline"],
             1 if request.form.get("featured") else 0), commit=True)
        flash(f"Job '{v['title']}' posted!","success")
        return redirect(url_for("company_dashboard"))
    return render_template("post_job.html", comp=comp, candidate=None, company=comp)

@app.route("/company/jobs/<int:job_id>/applicants")
@login_required_company
def job_applicants(job_id):
    comp = current_company()
    job = query("SELECT * FROM jobs WHERE id=? AND company_id=?", (job_id,comp["id"]), one=True)
    if not job: flash("Not found.","danger"); return redirect(url_for("company_dashboard"))
    return render_template("job_applicants.html", job=job, ranked=rank_applications(job_id),
        comp=comp, candidate=None, company=comp)

@app.route("/company/ranking/<int:job_id>")
@login_required_company
def ai_ranking(job_id):
    comp = current_company()
    job = query("SELECT * FROM jobs WHERE id=? AND company_id=?", (job_id,comp["id"]), one=True)
    if not job: flash("Not found.","danger"); return redirect(url_for("company_dashboard"))
    return render_template("ai_ranking.html", job=job, ranked=rank_applications(job_id),
        comp=comp, candidate=None, company=comp)

@app.route("/company/application/<int:app_id>/stage", methods=["POST"])
@login_required_company
def update_stage(app_id):
    stage = request.form.get("stage")
    if stage not in ["Applied","Shortlisted","Interview","Offer","Hired","Rejected"]:
        return jsonify({"error":"Invalid stage"}), 400
    query("UPDATE applications SET stage=? WHERE id=?", (stage,app_id), commit=True)
    return jsonify({"ok":True,"stage":stage})

@app.route("/company/jobs/<int:job_id>/toggle-status", methods=["POST"])
@login_required_company
def toggle_job_status(job_id):
    comp = current_company()
    job = query("SELECT * FROM jobs WHERE id=? AND company_id=?", (job_id,comp["id"]), one=True)
    if not job: return jsonify({"error":"Not found"}), 404
    new_status = "Closed" if job["status"]=="Active" else "Active"
    query("UPDATE jobs SET status=? WHERE id=?", (new_status,job_id), commit=True)
    return jsonify({"status":new_status})

@app.route("/company/profile", methods=["GET","POST"])
@login_required_company
def company_profile():
    comp = current_company()
    if request.method == "POST":
        f = ["name","industry","size","location","website","about"]
        v = {x: request.form.get(x,"").strip() for x in f}
        query("UPDATE companies SET name=?,industry=?,size=?,location=?,website=?,about=? WHERE id=?",
            (v["name"],v["industry"],v["size"],v["location"],v["website"],v["about"],comp["id"]), commit=True)
        session["user_name"] = v["name"]
        flash("Profile updated!","success"); return redirect(url_for("company_profile"))
    return render_template("company_profile.html", comp=comp, candidate=None, company=comp)

# ── Init & run ─────────────────────────────────────────────────────────────
with app.app_context():
    init_db()

if __name__ == "__main__":
    port  = int(os.environ.get("PORT", 5000))
    debug = os.environ.get("FLASK_ENV") != "production"
    app.run(debug=debug, host="0.0.0.0", port=port)
