#!/usr/bin/env python3
"""Exam Bill Calculator 1.2 — DUET undergraduate examination bill workflow.

Run: streamlit run exam_bill_calculator_1_2.py
Temporary mode: registration accepts any syntactically valid email and does not require email verification. Restrict access before production use.
The revised remuneration rates are seeded as draft/unapproved because formal source-page
metadata and scope interpretations must be confirmed by an authorized DUET administrator.
"""
from __future__ import annotations

import os, re, io, json, hmac, hashlib, secrets, smtplib, sqlite3, logging, math
from datetime import datetime, date, timedelta, timezone
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation
from email.message import EmailMessage
from pathlib import Path
from typing import Any, Optional, Iterable

import pandas as pd
import streamlit as st

APP_TITLE = "Exam Bill Calculator 1.2"
DEVELOPER = "Developed by Md. Zakirul Hasan | Lecturer | Department of Food Engineering | Dhaka University of Engineering & Technology (DUET), Gazipur"
DB_PATH = os.getenv("EXAM_BILL_DB", "exam_bill_calculator.sqlite3")
UPLOAD_LIMIT_MB = int(os.getenv("EXAM_BILL_UPLOAD_LIMIT_MB", "30"))
CURRENCY = "BDT"
CENT = Decimal("0.01")
BILL_DEPARTMENTS = ["Civil Engineering (CE)", "Electrical and Electronic Engineering (EEE)", "Mechanical Engineering (ME)", "Computer Science and Engineering (CSE)", "Textile Engineering (TE)", "Industrial & Production Engineering (IPE)", "Architecture", "Food Engineering (FE)", "Chemical Engineering (ChE)", "Materials and Metallurgical Engineering (MME)"]
TEACHING_DEPARTMENTS = BILL_DEPARTMENTS + ["Chemistry", "Mathematics", "Physics", "Humanities & Social Sciences"]
ROLES = ["Course Coordinator", "Department Administrator", "Authorized Reviewer/Approver", "System Administrator"]
STATUSES = ["Draft", "Under Review", "Approved", "Closed"]
CATEGORIES = ["Regular", "Review", "Special"]

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("exam_bill_calculator")

# Official rate inputs transcribed from the supplied master prompt. The rates remain
# unapproved until an authorized administrator records official source/page confirmation.
RATE_SEED = [
("Question paper setting",3600,"per question paper per section","individual","Undergraduate"),
("Question-paper moderation",2000,"per moderator per year-semester; total fee and cap interpretation requires verification","equal sharing","Undergraduate"),
("Question-paper typing",120,"per question paper","equal sharing","Undergraduate"),
("Comparison, sketching, correction and related question-paper work",1450,"per question paper","equal sharing","Undergraduate"),
("Question-paper printing",50,"per question paper","equal sharing","Undergraduate"),
("Answer-script examination",100,"per answer script per section; minimum 1000 per examiner per section","individual","Undergraduate"),
("Assignment assessment",50,"per course per teacher per student","individual","Undergraduate"),
("Class test",60,"per class test per participating student","individual","Undergraduate"),
("Answer-script scrutiny",20,"per answer script per semester; minimum 1500 per scrutinizer per year-semester; max two scrutinizers","individual","Undergraduate"),
("Course-file preparation",6000,"per course file","equal sharing","Undergraduate"),
("Survey work (practical subjects)",1000,"per course per student","equal sharing","Undergraduate"),
("Sessional (practical subjects)",320,"per course per student; minimum 1600 per teacher per course","equal sharing","Undergraduate"),
("Sessional (Architecture special rate)",115,"per course per student per contact hour; minimum 1600 per teacher per course","equal sharing","Undergraduate"),
("Central viva (sessional/practical)",150,"per semester per student","board sharing","Undergraduate"),
("Industrial attachment/training",750,"per student","equal sharing","Undergraduate"),
("Thesis/project examination",3000,"per thesis/project per examiner","individual","Undergraduate"),
("Preliminary oral examination",100,"per examinee per examiner; year/semester and board scope requires confirmation","board sharing","Undergraduate"),
("Final oral examination",250,"per examinee per examiner; year/semester and board scope requires confirmation","board sharing","Undergraduate"),
("Thesis/project supervision",6000,"per thesis/project; maximum 24000 per supervisor","individual","Undergraduate"),
("Capstone project",3000,"per student","equal sharing","Undergraduate"),
("Chief invigilator",1200,"per hour","individual","Undergraduate"),
("Additional chief invigilator",1100,"per hour","individual","Undergraduate"),
("Other invigilator",1000,"per hour","individual","Undergraduate"),
("Controller of examinations",1000,"per hour","individual","Undergraduate"),
("Deputy controller of examinations",900,"per shift","individual","Undergraduate"),
("Programmer/assistant controller of examinations",800,"per shift","individual","Undergraduate"),
("Assistant officer to chief invigilator",700,"per shift","individual","Undergraduate"),
("Other officer",550,"per shift","individual","Undergraduate"),
("Assistant/third-class employee",200,"per hour","individual","Undergraduate"),
("Fourth-class employee",175,"per hour","individual","Undergraduate"),
("Grade-sheet preparation — theoretical",60,"per course per student","equal sharing","Undergraduate"),
("Grade-sheet preparation — practical",50,"per course per student","equal sharing","Undergraduate"),
("Grade-sheet scrutiny — theoretical",30,"per course per student","equal sharing","Undergraduate"),
("Grade-sheet scrutiny — practical",12,"per course per student","equal sharing","Undergraduate"),
("Tabulation",90,"per student","equal sharing","Undergraduate"),
("Advisory fee",250,"per semester per student","individual","Undergraduate"),
("Verification of final graduation result",800,"per student","individual","Undergraduate"),
("Computerized result preparation",12,"per course per student","equal sharing","Undergraduate"),
("Student-wise grade and GPA list verification",24,"per student","equal sharing","Undergraduate"),
("Chairman, examination committee honorarium",4800,"per semester; category eligibility requires confirmation","individual","Undergraduate"),
("Course coordinator",3600,"per year per semester; Review included with Regular; Special Examination exception","individual","Undergraduate"),
]


def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")

def money(value: Any) -> Decimal:
    try:
        d = Decimal(str(value).strip())
        if not d.is_finite(): raise InvalidOperation
        return d
    except (InvalidOperation, ValueError, TypeError, AttributeError):
        raise ValueError("Enter a valid finite numeric amount or quantity.")

def bd_money(value: Any) -> str:
    return f"Tk. {money(value).quantize(CENT, rounding=ROUND_HALF_UP):,.2f}"

def safe_excel_text(value: Any) -> Any:
    if isinstance(value, str) and value.lstrip().startswith(("=", "+", "-", "@")):
        return "'" + value
    return value

def password_hash(password: str, salt: Optional[bytes] = None) -> tuple[str, str]:
    salt = salt or secrets.token_bytes(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode(), salt, 310_000)
    return salt.hex(), digest.hex()

def password_matches(password: str, salt_hex: str, hash_hex: str) -> bool:
    _, candidate = password_hash(password, bytes.fromhex(salt_hex))
    return hmac.compare_digest(candidate, hash_hex)

def config_value(name: str, default: str = "") -> str:
    """Read configuration from environment variables or Streamlit secrets."""
    value = os.getenv(name)
    if value:
        return value
    try:
        value = st.secrets.get(name, default)
        return str(value) if value is not None else default
    except Exception:
        return default

SCHEMA = [
"""CREATE TABLE IF NOT EXISTS departments(id INTEGER PRIMARY KEY, name TEXT UNIQUE NOT NULL, bill_owner INTEGER NOT NULL DEFAULT 0, active INTEGER NOT NULL DEFAULT 1)""",
"""CREATE TABLE IF NOT EXISTS users(id INTEGER PRIMARY KEY, email TEXT UNIQUE NOT NULL COLLATE NOCASE, full_name TEXT NOT NULL, salt TEXT NOT NULL, password_hash TEXT NOT NULL, role TEXT NOT NULL, department_id INTEGER, verified INTEGER NOT NULL DEFAULT 0, active INTEGER NOT NULL DEFAULT 1, failed_logins INTEGER NOT NULL DEFAULT 0, locked_until TEXT, created_at TEXT NOT NULL, FOREIGN KEY(department_id) REFERENCES departments(id))""",
"""CREATE TABLE IF NOT EXISTS email_tokens(id INTEGER PRIMARY KEY, user_id INTEGER NOT NULL, token_hash TEXT NOT NULL, expires_at TEXT NOT NULL, used INTEGER NOT NULL DEFAULT 0, FOREIGN KEY(user_id) REFERENCES users(id))""",
"""CREATE TABLE IF NOT EXISTS exam_sessions(id INTEGER PRIMARY KEY, bill_department_id INTEGER NOT NULL, academic_session TEXT NOT NULL, exam_name TEXT NOT NULL, exam_year INTEGER NOT NULL, semester INTEGER NOT NULL, exam_type TEXT NOT NULL, start_date TEXT, end_date TEXT, preparation_date TEXT, status TEXT NOT NULL DEFAULT 'Draft', reference_no TEXT, remarks TEXT, created_by INTEGER NOT NULL, created_at TEXT NOT NULL, updated_at TEXT NOT NULL, FOREIGN KEY(bill_department_id) REFERENCES departments(id), FOREIGN KEY(created_by) REFERENCES users(id))""",
"""CREATE TABLE IF NOT EXISTS class_groups(id INTEGER PRIMARY KEY, session_id INTEGER NOT NULL, academic_year INTEGER NOT NULL, semester INTEGER NOT NULL, category TEXT NOT NULL, student_count INTEGER, count_confirmed INTEGER NOT NULL DEFAULT 0, remarks TEXT, created_at TEXT NOT NULL, UNIQUE(session_id,academic_year,semester,category), FOREIGN KEY(session_id) REFERENCES exam_sessions(id))""",
"""CREATE TABLE IF NOT EXISTS teachers(id INTEGER PRIMARY KEY, full_name TEXT NOT NULL, employee_id TEXT, home_department_id INTEGER, designation TEXT, email TEXT, active INTEGER NOT NULL DEFAULT 1, created_at TEXT NOT NULL, FOREIGN KEY(home_department_id) REFERENCES departments(id))""",
"""CREATE TABLE IF NOT EXISTS courses(id INTEGER PRIMARY KEY, code TEXT NOT NULL, title TEXT NOT NULL, teaching_department_id INTEGER NOT NULL, subject_type TEXT NOT NULL DEFAULT 'Theory', contact_hours TEXT, active INTEGER NOT NULL DEFAULT 1, UNIQUE(code,teaching_department_id), FOREIGN KEY(teaching_department_id) REFERENCES departments(id))""",
"""CREATE TABLE IF NOT EXISTS course_offerings(id INTEGER PRIMARY KEY, group_id INTEGER NOT NULL, course_id INTEGER NOT NULL, registration_count INTEGER, count_confirmed INTEGER NOT NULL DEFAULT 0, remarks TEXT, UNIQUE(group_id,course_id), FOREIGN KEY(group_id) REFERENCES class_groups(id), FOREIGN KEY(course_id) REFERENCES courses(id))""",
"""CREATE TABLE IF NOT EXISTS sources(id INTEGER PRIMARY KEY, title TEXT NOT NULL, document_ref TEXT, effective_date TEXT, file_name TEXT, remarks TEXT, created_at TEXT NOT NULL)""",
"""CREATE TABLE IF NOT EXISTS rates(id INTEGER PRIMARY KEY, item_name TEXT NOT NULL, source_wording TEXT, description TEXT, amount TEXT NOT NULL, currency TEXT NOT NULL DEFAULT 'BDT', unit TEXT NOT NULL, minimum_amount TEXT, maximum_amount TEXT, min_scope TEXT, max_scope TEXT, distribution TEXT NOT NULL, study_level TEXT NOT NULL DEFAULT 'Undergraduate', applicability TEXT NOT NULL DEFAULT 'REQUIRES AUTHORIZED VERIFICATION', subject_type TEXT, effective_date TEXT, expiry_date TEXT, source_id INTEGER, source_page TEXT, source_item TEXT, approval_status TEXT NOT NULL DEFAULT 'Draft', version INTEGER NOT NULL DEFAULT 1, change_reason TEXT, created_by INTEGER, approved_by INTEGER, created_at TEXT NOT NULL, active INTEGER NOT NULL DEFAULT 0, FOREIGN KEY(source_id) REFERENCES sources(id), FOREIGN KEY(created_by) REFERENCES users(id), FOREIGN KEY(approved_by) REFERENCES users(id))""",
"""CREATE TABLE IF NOT EXISTS duty_assignments(id INTEGER PRIMARY KEY, session_id INTEGER NOT NULL, group_id INTEGER, offering_id INTEGER, teacher_id INTEGER NOT NULL, rate_id INTEGER NOT NULL, item_name TEXT NOT NULL, quantity TEXT NOT NULL, unit TEXT NOT NULL, eligible_assignees TEXT NOT NULL DEFAULT '[]', formula TEXT NOT NULL, scope TEXT NOT NULL, status TEXT NOT NULL DEFAULT 'Draft', remarks TEXT, source_ref TEXT, created_by INTEGER NOT NULL, created_at TEXT NOT NULL, updated_at TEXT NOT NULL, FOREIGN KEY(session_id) REFERENCES exam_sessions(id), FOREIGN KEY(group_id) REFERENCES class_groups(id), FOREIGN KEY(offering_id) REFERENCES course_offerings(id), FOREIGN KEY(teacher_id) REFERENCES teachers(id), FOREIGN KEY(rate_id) REFERENCES rates(id), FOREIGN KEY(created_by) REFERENCES users(id))""",
"""CREATE TABLE IF NOT EXISTS bill_headers(id INTEGER PRIMARY KEY, session_id INTEGER NOT NULL, teacher_id INTEGER NOT NULL, reference_no TEXT UNIQUE NOT NULL, status TEXT NOT NULL DEFAULT 'Draft', total TEXT NOT NULL DEFAULT '0', created_at TEXT NOT NULL, updated_at TEXT NOT NULL, FOREIGN KEY(session_id) REFERENCES exam_sessions(id), FOREIGN KEY(teacher_id) REFERENCES teachers(id))""",
"""CREATE TABLE IF NOT EXISTS bill_lines(id INTEGER PRIMARY KEY, bill_id INTEGER NOT NULL, duty_id INTEGER NOT NULL, item_name TEXT NOT NULL, quantity TEXT NOT NULL, unit TEXT NOT NULL, rate_amount TEXT NOT NULL, rate_version INTEGER NOT NULL, gross_amount TEXT NOT NULL, adjustment TEXT NOT NULL, distribution TEXT NOT NULL, payable_amount TEXT NOT NULL, breakdown TEXT NOT NULL, validation_status TEXT NOT NULL, snapshot_json TEXT NOT NULL, FOREIGN KEY(bill_id) REFERENCES bill_headers(id), FOREIGN KEY(duty_id) REFERENCES duty_assignments(id))""",
"""CREATE TABLE IF NOT EXISTS approvals(id INTEGER PRIMARY KEY, session_id INTEGER NOT NULL, actor_id INTEGER NOT NULL, action TEXT NOT NULL, from_status TEXT, to_status TEXT, reason TEXT, created_at TEXT NOT NULL, FOREIGN KEY(session_id) REFERENCES exam_sessions(id), FOREIGN KEY(actor_id) REFERENCES users(id))""",
"""CREATE TABLE IF NOT EXISTS import_batches(id INTEGER PRIMARY KEY, filename TEXT NOT NULL, imported_by INTEGER NOT NULL, status TEXT NOT NULL, summary TEXT, created_at TEXT NOT NULL, FOREIGN KEY(imported_by) REFERENCES users(id))""",
"""CREATE TABLE IF NOT EXISTS reconciliation(id INTEGER PRIMARY KEY, batch_id INTEGER NOT NULL, sheet_name TEXT, source_row TEXT, teacher TEXT, group_name TEXT, source_item TEXT, source_quantity TEXT, source_rate TEXT, source_amount TEXT, expected_rate TEXT, expected_amount TEXT, difference TEXT, validation_note TEXT, status TEXT NOT NULL DEFAULT 'Unresolved', reviewer_id INTEGER, decision_date TEXT, FOREIGN KEY(batch_id) REFERENCES import_batches(id), FOREIGN KEY(reviewer_id) REFERENCES users(id))""",
"""CREATE TABLE IF NOT EXISTS audit_logs(id INTEGER PRIMARY KEY, actor_id INTEGER, actor_email TEXT, action TEXT NOT NULL, entity TEXT NOT NULL, record_id TEXT, before_json TEXT, after_json TEXT, reason TEXT, created_at TEXT NOT NULL)""",
"""CREATE TABLE IF NOT EXISTS exports(id INTEGER PRIMARY KEY, user_id INTEGER NOT NULL, session_id INTEGER, export_type TEXT NOT NULL, filename TEXT NOT NULL, created_at TEXT NOT NULL, FOREIGN KEY(user_id) REFERENCES users(id), FOREIGN KEY(session_id) REFERENCES exam_sessions(id))""",
"CREATE INDEX IF NOT EXISTS idx_sessions_dept ON exam_sessions(bill_department_id,status)",
"CREATE INDEX IF NOT EXISTS idx_duties_session ON duty_assignments(session_id,teacher_id)",
"CREATE INDEX IF NOT EXISTS idx_audit_created ON audit_logs(created_at)",
]

class Repository:
    def __init__(self, path: str = DB_PATH):
        self.path = path
        Path(path).expanduser().resolve().parent.mkdir(parents=True, exist_ok=True)
        self.initialize()
    def connect(self):
        con = sqlite3.connect(self.path, timeout=30, isolation_level=None)
        con.row_factory = sqlite3.Row
        con.execute("PRAGMA foreign_keys=ON")
        con.execute("PRAGMA busy_timeout=30000")
        return con
    def initialize(self):
        with self.connect() as con:
            con.execute("PRAGMA journal_mode=WAL")
            for sql in SCHEMA: con.execute(sql)
            for name in TEACHING_DEPARTMENTS:
                con.execute("INSERT OR IGNORE INTO departments(name,bill_owner) VALUES(?,?)", (name, int(name in BILL_DEPARTMENTS)))
            # Temporary no-verification mode: existing accounts become usable without email verification.
            con.execute("UPDATE users SET verified=1 WHERE verified=0")
            # Seed rates as drafts, never approved/active without authorized source review.
            for name, amount, unit, distribution, level in RATE_SEED:
                exists = con.execute("SELECT 1 FROM rates WHERE item_name=? AND version=1", (name,)).fetchone()
                if not exists:
                    con.execute("INSERT INTO rates(item_name,source_wording,description,amount,unit,distribution,study_level,applicability,approval_status,version,change_reason,created_at,active) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,0)", (name,name,"Seeded from supplied master prompt; confirm exact official PDF page/item and scope.",str(amount),unit,distribution,level,"REQUIRES AUTHORIZED VERIFICATION","Draft",1,"Initial import; requires official source metadata confirmation",now_iso()))
    def audit(self, actor, action, entity, record_id=None, before=None, after=None, reason=None):
        with self.connect() as c:
            c.execute("INSERT INTO audit_logs(actor_id,actor_email,action,entity,record_id,before_json,after_json,reason,created_at) VALUES(?,?,?,?,?,?,?,?,?)",(actor["id"] if actor else None, actor["email"] if actor else None, action,entity,str(record_id) if record_id is not None else None,json.dumps(before,default=str,ensure_ascii=False) if before is not None else None,json.dumps(after,default=str,ensure_ascii=False) if after is not None else None,reason,now_iso()))
    def rows(self, sql, params=()):
        with self.connect() as c: return [dict(r) for r in c.execute(sql,params).fetchall()]
    def one(self, sql, params=()):
        with self.connect() as c:
            r=c.execute(sql,params).fetchone(); return dict(r) if r else None
    def execute(self, sql, params=()):
        with self.connect() as c:
            c.execute("BEGIN IMMEDIATE")
            try:
                cur=c.execute(sql,params); c.execute("COMMIT"); return cur.lastrowid
            except Exception:
                c.execute("ROLLBACK"); raise

repo = Repository()

# ---------- Security and authorization ----------
def require_role(user: dict, allowed: Iterable[str]):
    if not user or not user.get("active") or user.get("role") not in set(allowed):
        raise PermissionError("Your account does not have permission for this action.")

def send_verification_email(email: str, token: str) -> None:
    host=os.getenv("SMTP_HOST"); sender=os.getenv("SMTP_FROM"); password=os.getenv("SMTP_PASSWORD")
    if not all([host,sender,password]): raise RuntimeError("Email verification is unavailable: configure SMTP_HOST, SMTP_FROM, SMTP_PASSWORD and optionally SMTP_PORT/SMTP_USER before registration.")
    msg=EmailMessage(); msg["Subject"]="DUET Exam Bill Calculator email verification"; msg["From"]=sender; msg["To"]=email
    msg.set_content(f"Your verification token is: {token}\nIt expires in 30 minutes. If you did not request this, ignore this email.")
    with smtplib.SMTP(host,int(os.getenv("SMTP_PORT","587")),timeout=15) as server:
        server.starttls(); user=os.getenv("SMTP_USER") or sender; server.login(user,password); server.send_message(msg)

def register_user(email: str, full_name: str, password: str, role: str, department_id: Optional[int], bootstrap_token: str = "") -> str:
    """Create an account without email verification; a secret token can bootstrap the first administrator."""
    email = email.strip().lower()
    if not re.fullmatch(r"[^\s@]+@[^\s@]+\.[^\s@]+", email):
        raise ValueError("Enter a valid email address.")
    if not full_name.strip():
        raise ValueError("Enter your full name.")
    if len(password) < 12 or not re.search(r"[A-Za-z]", password) or not re.search(r"\d", password):
        raise ValueError("Password must be at least 12 characters and include letters and digits.")
    # Self-registration remains least-privileged unless the operator supplies the
    # deployment-only bootstrap secret. This is needed to approve the initial rates.
    expected = config_value("EXAM_BILL_ADMIN_BOOTSTRAP_TOKEN")
    assigned_role = "Course Coordinator"
    if expected and bootstrap_token and hmac.compare_digest(bootstrap_token, expected):
        if repo.one("SELECT id FROM users WHERE role='System Administrator' AND active=1"):
            raise PermissionError("A System Administrator already exists; bootstrap registration is disabled.")
        assigned_role = "System Administrator"
    elif bootstrap_token:
        raise PermissionError("Administrator bootstrap token is incorrect or not configured.")
    salt, hashed = password_hash(password)
    uid = repo.execute(
        "INSERT INTO users(email,full_name,salt,password_hash,role,department_id,verified,active,created_at) VALUES(?,?,?,?,?,?,1,1,?)",
        (email, full_name.strip(), salt, hashed, assigned_role, department_id, now_iso()),
    )
    repo.audit(None, "REGISTER_NO_EMAIL_VERIFICATION", "user", uid,
               after={"email": email, "role": assigned_role, "department_id": department_id},
               reason="Temporary open-registration mode; email ownership was not verified.")
    return assigned_role

def login_user(email: str, password: str):
    email=email.strip().lower(); u=repo.one("SELECT * FROM users WHERE email=?",(email,))
    if not u: raise ValueError("Invalid email or password.")
    if not u["active"]: raise PermissionError("This account is inactive.")
    if u["locked_until"] and u["locked_until"] > now_iso(): raise PermissionError("Account temporarily locked after failed attempts.")
    if not password_matches(password,u["salt"],u["password_hash"]):
        fails=int(u["failed_logins"])+1; locked=(datetime.now(timezone.utc)+timedelta(minutes=15)).isoformat() if fails>=5 else None
        repo.execute("UPDATE users SET failed_logins=?,locked_until=? WHERE id=?",(fails,locked,u["id"]))
        repo.audit(None,"LOGIN_FAILED","user",u["id"])
        raise ValueError("Invalid email or password.")
    repo.execute("UPDATE users SET failed_logins=0,locked_until=NULL WHERE id=?",(u["id"],))
    repo.audit(u,"LOGIN","user",u["id"])
    return {k:u[k] for k in ("id","email","full_name","role","department_id","verified","active")}

def visible_sessions(user):
    if user["role"]=="System Administrator":
        return repo.rows("SELECT s.*,d.name department FROM exam_sessions s JOIN departments d ON d.id=s.bill_department_id ORDER BY s.id DESC")
    if user["department_id"]:
        return repo.rows("SELECT s.*,d.name department FROM exam_sessions s JOIN departments d ON d.id=s.bill_department_id WHERE s.bill_department_id=? ORDER BY s.id DESC",(user["department_id"],))
    return []

def check_session_access(user, session_id: int):
    s=repo.one("SELECT * FROM exam_sessions WHERE id=?",(session_id,))
    if not s: raise ValueError("Examination session not found.")
    if user["role"]!="System Administrator" and user.get("department_id")!=s["bill_department_id"]:
        raise PermissionError("You are not authorized for this department's examination session.")
    return s

# ---------- Calculation engine: no Streamlit dependency ----------
def calculate_line(*, quantity: Any, rate: Any, unit: str, item_name: str, assignee_count: int = 1,
                   minimum: Any = None, maximum: Any = None, minimum_scope: str = "", maximum_scope: str = "",
                   distribution: str = "individual", examiner_script_minimum: bool = False,
                   confirmed: bool = False, rate_approved: bool = False, applicability: str = "") -> dict:
    """Calculate one eligible line transparently; caller must enforce wider aggregate scopes."""
    q=money(quantity); r=money(rate)
    if q<0 or r<0: raise ValueError("Quantity and rate cannot be negative.")
    if q==0: raise ValueError("Quantity must be positive; missing quantities must not be silently treated as zero.")
    if assignee_count<1: raise ValueError("At least one eligible assignee is required.")
    gross=q*r; adjustment=Decimal("0")
    if minimum is not None and str(minimum)!="":
        mn=money(minimum)
        if mn<0: raise ValueError("Minimum cannot be negative.")
        if gross<mn: adjustment=mn-gross; gross=mn
    if maximum is not None and str(maximum)!="":
        mx=money(maximum)
        if mx<0: raise ValueError("Maximum cannot be negative.")
        if gross>mx: adjustment=mx-gross; gross=mx
    if examiner_script_minimum and gross<Decimal("1000"):
        adjustment=Decimal("1000")-gross; gross=Decimal("1000")
    if distribution in ("equal sharing","board sharing"):
        per_person=(gross/Decimal(assignee_count)).quantize(CENT,rounding=ROUND_HALF_UP)
    elif distribution=="individual": per_person=gross.quantize(CENT,rounding=ROUND_HALF_UP)
    else: raise ValueError("Unknown distribution method; requires authorized configuration.")
    warnings=[]
    if not confirmed: warnings.append("Quantity is not explicitly confirmed.")
    if not rate_approved: warnings.append("Rate is not approved/active.")
    if applicability and applicability!="Approved": warnings.append("Applicability requires authorized verification.")
    return {"quantity":str(q),"rate":str(r),"unit":unit,"gross_before_adjustment":str(q*r),"adjustment":str(adjustment),"gross_after_adjustment":str(gross),"distribution":distribution,"assignee_count":assignee_count,"per_person":str(per_person),"total_allocated":str(per_person*assignee_count if distribution!="individual" else per_person),"warnings":warnings,"blocking":bool(warnings)}

def validate_group_counts(groups: list[dict]) -> list[str]:
    issues=[]
    seen=set()
    for g in groups:
        key=(g.get("session_id"),g.get("academic_year"),g.get("semester"),g.get("category"))
        if key in seen: issues.append(f"Duplicate class group: {key}")
        seen.add(key)
        n=g.get("student_count")
        if n is None or int(n)<0: issues.append(f"Group {key}: student count missing or negative.")
        if not g.get("count_confirmed"): issues.append(f"Group {key}: student count not confirmed.")
    return issues

def validate_scrutinizers(year: int, semester: int, teacher_ids: list[int]) -> list[str]:
    return [f"Year {year}, semester {semester}: at most two answer-script scrutinizers are allowed." ] if len(set(teacher_ids))>2 else []

def generate_teacher_bill(session_id: int, teacher_id: int) -> dict:
    """Persist an auditable bill snapshot; only approved rates are calculated into payable lines."""
    session=repo.one("SELECT * FROM exam_sessions WHERE id=?",(session_id,))
    teacher=repo.one("SELECT * FROM teachers WHERE id=?",(teacher_id,))
    if not session or not teacher: raise ValueError("Session or teacher not found.")
    duties=repo.rows("SELECT d.*,r.amount rate_amount,r.version rate_version,r.approval_status,r.active rate_active,r.applicability,r.minimum_amount,r.maximum_amount,r.min_scope,r.max_scope,r.distribution rate_distribution FROM duty_assignments d JOIN rates r ON r.id=d.rate_id WHERE d.session_id=? AND d.teacher_id=? ORDER BY d.id",(session_id,teacher_id))
    ref=f"EBC-{session_id:05d}-{teacher_id:05d}-{datetime.now().strftime('%Y%m%d%H%M%S')}"
    with repo.connect() as c:
        c.execute("BEGIN IMMEDIATE")
        try:
            bid=c.execute("INSERT INTO bill_headers(session_id,teacher_id,reference_no,status,total,created_at,updated_at) VALUES(?,?,?,'Draft','0',?,?)",(session_id,teacher_id,ref,now_iso(),now_iso())).lastrowid
            total=Decimal("0"); issues=[]
            for d in duties:
                assignees=json.loads(d["eligible_assignees"] or "[]")
                if not assignees: assignees=[teacher_id]
                isapproved=d["approval_status"]=="Approved" and bool(d["rate_active"])
                result=calculate_line(quantity=d["quantity"],rate=d["rate_amount"],unit=d["unit"],item_name=d["item_name"],assignee_count=len(set(assignees)),minimum=d["minimum_amount"],maximum=d["maximum_amount"],minimum_scope=d["min_scope"] or "",maximum_scope=d["max_scope"] or "",distribution=d["rate_distribution"],examiner_script_minimum=d["item_name"]=="Answer-script examination",confirmed=True,rate_approved=isapproved,applicability=d["applicability"])
                if result["blocking"]: issues.extend([f"Duty {d['id']}: {w}" for w in result["warnings"]]); payable=Decimal("0"); status="BLOCKED"
                else: payable=money(result["per_person"]); total+=payable; status="VALID"
                snap={"session_id":session_id,"teacher_id":teacher_id,"duty_id":d["id"],"rate_id":d["rate_id"],"rate_version":d["rate_version"],"calculation":result}
                c.execute("INSERT INTO bill_lines(bill_id,duty_id,item_name,quantity,unit,rate_amount,rate_version,gross_amount,adjustment,distribution,payable_amount,breakdown,validation_status,snapshot_json) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)",(bid,d["id"],d["item_name"],d["quantity"],d["unit"],d["rate_amount"],d["rate_version"],result["gross_after_adjustment"],result["adjustment"],d["rate_distribution"],str(payable),json.dumps(result),status,json.dumps(snap)))
            c.execute("UPDATE bill_headers SET total=?,updated_at=? WHERE id=?",(str(total.quantize(CENT)),now_iso(),bid)); c.execute("COMMIT")
        except Exception: c.execute("ROLLBACK"); raise
    return {"id":bid,"reference_no":ref,"total":str(total.quantize(CENT)),"issues":issues,"line_count":len(duties)}

# ---------- Import/export ----------
def inspect_excel(uploaded_file) -> tuple[dict[str,pd.DataFrame], list[str]]:
    name=uploaded_file.name.lower()
    if not name.endswith((".xlsx",".xlsm")): raise ValueError("Only .xlsx and .xlsm files are supported.")
    data=uploaded_file.getvalue()
    if len(data)>UPLOAD_LIMIT_MB*1024*1024: raise ValueError(f"File exceeds the {UPLOAD_LIMIT_MB} MB upload limit.")
    xl=pd.ExcelFile(io.BytesIO(data),engine="openpyxl")
    sheets={}
    notes=[]
    for sheet in xl.sheet_names:
        try:
            # data_only=True reads cached formula results; a second pass identifies formulas.
            df=pd.read_excel(io.BytesIO(data),sheet_name=sheet,header=None,engine="openpyxl")
            sheets[sheet]=df
            wb=__import__('openpyxl').load_workbook(io.BytesIO(data),read_only=True,data_only=False,keep_vba=name.endswith('.xlsm'))
            ws=wb[sheet]; formula_count=sum(1 for row in ws.iter_rows() for cell in row if isinstance(cell.value,str) and cell.value.startswith('='))
            if formula_count: notes.append(f"{sheet}: {formula_count} formulas detected; formulas are not executed and must be reconciled.")
            wb.close()
        except Exception as exc: notes.append(f"Could not fully inspect sheet {sheet}: {exc}")
    return sheets,notes

def workbook_bytes(df: pd.DataFrame, sheet_name: str = "Report", title: str = APP_TITLE) -> bytes:
    out=io.BytesIO()
    safe=df.copy()
    for col in safe.columns: safe[col]=safe[col].map(safe_excel_text)
    with pd.ExcelWriter(out,engine="openpyxl") as writer:
        safe.to_excel(writer,index=False,sheet_name=sheet_name[:31],startrow=3)
        ws=writer.sheets[sheet_name[:31]]
        ws.cell(1,1,title); ws.cell(2,1,f"Generated {now_iso()} | Draft report unless explicitly approved")
        ws.freeze_panes="A5"; ws.auto_filter.ref=ws.dimensions
        from openpyxl.styles import Font, PatternFill, Alignment
        ws.cell(1,1).font=Font(bold=True,size=14); ws.cell(2,1).font=Font(italic=True,size=9)
        for cell in ws[4]: cell.font=Font(bold=True,color="FFFFFF"); cell.fill=PatternFill("solid",fgColor="24476B"); cell.alignment=Alignment(wrap_text=True)
        for col in ws.columns:
            letter=col[0].column_letter; maxlen=max((len(str(x.value or "")) for x in col),default=10)
            ws.column_dimensions[letter].width=min(max(maxlen+2,12),42)
    return out.getvalue()

# ---------- UI helpers ----------
def log_ui(user, action, entity, rid=None, before=None, after=None, reason=None):
    repo.audit(user,action,entity,rid,before,after,reason)
def department_options(bill_only=False):
    rows=repo.rows("SELECT id,name FROM departments WHERE active=1" + (" AND bill_owner=1" if bill_only else "") + " ORDER BY name")
    return rows

def session_label(s): return f"#{s['id']} — {s['department']} — {s['exam_name']} ({s['academic_session']}, {s['exam_year']}, {s['semester']}-{s['exam_type']}) [{s['status']}]"
def can_edit_session(user,s): return user["role"]=="System Administrator" or user.get("department_id")==s["bill_department_id"]
def show_table(rows, key=None):
    if rows: st.dataframe(pd.DataFrame(rows),use_container_width=True,hide_index=True)
    else: st.info("No records yet.")

def page_login():
    st.title(APP_TITLE); st.caption("DUET undergraduate examination bill preparation, review and audit")
    login_tab, reg_tab=st.tabs(["Sign in","Register"])
    with login_tab:
        with st.form("login_form"):
            email=st.text_input("Email address",placeholder="you@example.com"); pw=st.text_input("Password",type="password"); ok=st.form_submit_button("Sign in",use_container_width=True)
        if ok:
            try: st.session_state.user=login_user(email,pw); st.rerun()
            except Exception as e: st.error(str(e))
    with reg_tab:
        st.warning("Temporary open-registration mode: any valid email address is accepted without verification. Do not use this setting for a public production deployment; email ownership and institutional affiliation are not verified.")
        deps=department_options()
        with st.form("register_form"):
            full=st.text_input("Full name"); remail=st.text_input("Email address",key="reg_email"); rpw=st.text_input("Create password (12+ characters)",type="password",key="reg_pw")
            st.caption("New accounts receive Course Coordinator access. To initialize the app administrator, enter the deployment bootstrap token configured by the app owner.")
            dep=st.selectbox("Department access scope",["Not assigned"]+[f"{d['id']} — {d['name']}" for d in deps])
            bootstrap=st.text_input("Administrator bootstrap token (optional)",type="password",help="Leave blank for a standard account. Only works when EXAM_BILL_ADMIN_BOOTSTRAP_TOKEN is configured in Streamlit secrets.")
            submitted=st.form_submit_button("Register")
        if submitted:
            try:
                assigned_role=register_user(remail,full,rpw,"Course Coordinator",None if dep=="Not assigned" else int(dep.split(" — ")[0]),bootstrap)
                st.success(f"Account created as {assigned_role}. You can sign in immediately; email verification is disabled in this temporary mode.")
            except sqlite3.IntegrityError: st.error("An account with this email already exists.")
            except Exception as e: st.error(str(e))

def page_dashboard(user):
    st.header("Dashboard")
    sessions=visible_sessions(user); ids=[s["id"] for s in sessions]
    bills=repo.rows("SELECT b.*,t.full_name teacher,s.exam_name FROM bill_headers b JOIN teachers t ON t.id=b.teacher_id JOIN exam_sessions s ON s.id=b.session_id WHERE b.session_id IN ("+",".join("?" for _ in ids)+") ORDER BY b.id DESC",ids) if ids else []
    rates=repo.rows("SELECT approval_status,COUNT(*) n FROM rates GROUP BY approval_status")
    unresolved=repo.one("SELECT COUNT(*) n FROM reconciliation WHERE status='Unresolved'")["n"]
    cols=st.columns(4); cols[0].metric("Accessible sessions",len(sessions)); cols[1].metric("Draft bills",sum(b["status"]=="Draft" for b in bills)); cols[2].metric("Approved bills",sum(b["status"]=="Approved" for b in bills)); cols[3].metric("Unresolved reconciliation rows",unresolved)
    st.subheader("Recent examination sessions"); show_table(sessions[:10])
    st.subheader("Rate approval status"); show_table(rates)
    st.caption("Only records in your authorized department scope are shown. Historical archive counts are not reused as current quantities.")

def page_sessions(user):
    st.header("Examination sessions")
    deps=department_options(bill_only=True)
    if user["role"] in ("System Administrator","Department Administrator","Course Coordinator"):
        with st.expander("Create examination session",expanded=False):
            with st.form("create_session"):
                dept=st.selectbox("Bill for which department?",deps,format_func=lambda d:d["name"])
                acad=st.text_input("Academic session / batch",placeholder="2024-2025"); name=st.text_input("Examination name",placeholder="1st semester examination")
                yr=st.number_input("Examination year",min_value=2000,max_value=2100,value=date.today().year); sem=st.selectbox("Semester",[1,2]); typ=st.selectbox("Examination category",CATEGORIES)
                c1,c2=st.columns(2); start=c1.date_input("Start date",value=None); end=c2.date_input("End date",value=None)
                prep=st.date_input("Bill preparation date",value=date.today()); ref=st.text_input("Reference number"); remarks=st.text_area("Remarks")
                create=st.form_submit_button("Create session")
            if create:
                if not acad.strip() or not name.strip(): st.error("Academic session and examination name are required.")
                elif start and end and start>end: st.error("End date cannot precede start date.")
                elif user["role"]!="System Administrator" and user.get("department_id")!=dept["id"]: st.error("You may create sessions only for your authorized department.")
                else:
                    try:
                        sid=repo.execute("INSERT INTO exam_sessions(bill_department_id,academic_session,exam_name,exam_year,semester,exam_type,start_date,end_date,preparation_date,status,reference_no,remarks,created_by,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,'Draft',?,?,?,?,?)",(dept["id"],acad.strip(),name.strip(),int(yr),int(sem),typ,start.isoformat() if start else None,end.isoformat() if end else None,prep.isoformat(),ref.strip() or None,remarks.strip() or None,user["id"],now_iso(),now_iso()))
                        log_ui(user,"CREATE","exam_session",sid,after={"department":dept["name"],"exam_name":name,"academic_session":acad}); st.success(f"Created session #{sid}."); st.rerun()
                    except Exception as e: st.error(f"Could not create session: {e}")
    rows=visible_sessions(user); show_table(rows)
    if rows:
        selected=st.selectbox("Manage session",rows,format_func=session_label,key="session_manage")
        if can_edit_session(user,selected):
            st.subheader("Class-examination groups")
            groups=repo.rows("SELECT * FROM class_groups WHERE session_id=? ORDER BY academic_year,semester,category",(selected["id"],)); show_table(groups)
            if selected["status"] not in ("Approved","Closed") and user["role"] in ("System Administrator","Department Administrator","Course Coordinator"):
                with st.form("create_group"):
                    year=st.selectbox("Academic year",[1,2,3,4]); sem=st.selectbox("Class semester",[1,2],key="group_sem"); cat=st.selectbox("Group category",CATEGORIES,key="group_cat")
                    count=st.number_input("Explicit class student count",min_value=0,step=1,value=0); confirmed=st.checkbox("I confirm this student count from an authorized source")
                    gre=st.text_area("Group remarks"); add=st.form_submit_button("Add class group")
                if add:
                    try:
                        gid=repo.execute("INSERT INTO class_groups(session_id,academic_year,semester,category,student_count,count_confirmed,remarks,created_at) VALUES(?,?,?,?,?,?,?,?)",(selected["id"],year,sem,cat,int(count),int(confirmed),gre.strip() or None,now_iso()))
                        log_ui(user,"CREATE","class_group",gid,after={"session_id":selected["id"],"year":year,"semester":sem,"category":cat,"student_count":count,"confirmed":confirmed}); st.success("Class group added."); st.rerun()
                    except sqlite3.IntegrityError: st.error("This year-semester-category group already exists in this session.")
                    except Exception as e: st.error(str(e))
            st.subheader("Session workflow")
            if selected["status"] not in ("Approved","Closed"):
                with st.form("session_workflow"):
                    target=st.selectbox("Action",["Submit for review","Return to draft","Approve session","Close session"])
                    reason=st.text_input("Reason / review note",required=True); act=st.form_submit_button("Record workflow action")
                if act:
                    try:
                        allowed={"Submit for review":("Draft","Under Review"),"Return to draft":("Under Review","Draft"),"Approve session":("Under Review","Approved"),"Close session":("Approved","Closed")}
                        old,new=allowed[target]
                        if selected["status"]!=old: raise ValueError(f"Action requires status {old}; current status is {selected['status']}.")
                        if target=="Approve session": require_role(user,["Authorized Reviewer/Approver","System Administrator"])
                        if target=="Close session": require_role(user,["Department Administrator","System Administrator"])
                        if target=="Approve session" and repo.one("SELECT COUNT(*) n FROM rates WHERE active=1 AND approval_status='Approved'")["n"]==0: raise ValueError("No approved active rates are configured.")
                        repo.execute("UPDATE exam_sessions SET status=?,updated_at=? WHERE id=?",(new,now_iso(),selected["id"]))
                        repo.execute("INSERT INTO approvals(session_id,actor_id,action,from_status,to_status,reason,created_at) VALUES(?,?,?,?,?,?,?)",(selected["id"],user["id"],target,old,new,reason,now_iso()))
                        log_ui(user,target,"exam_session",selected["id"],before={"status":old},after={"status":new},reason=reason); st.success(f"Session status changed to {new}."); st.rerun()
                    except Exception as e: st.error(str(e))

def page_departments_teachers_courses(user):
    st.header("Departments, teachers and courses")
    tabs=st.tabs(["Departments","Teachers","Courses / offerings"])
    with tabs[0]:
        show_table(department_options())
        if user["role"]=="System Administrator":
            with st.form("new_dept"):
                dn=st.text_input("Additional teaching department"); bill=st.checkbox("Eligible to own undergraduate bills (authorized degree-awarding department only)"); save=st.form_submit_button("Add department")
            if save:
                if not dn.strip(): st.error("Department name required.")
                else:
                    try: did=repo.execute("INSERT INTO departments(name,bill_owner) VALUES(?,?)",(dn.strip(),int(bill))); log_ui(user,"CREATE","department",did,after={"name":dn,"bill_owner":bill}); st.success("Department added."); st.rerun()
                    except sqlite3.IntegrityError: st.error("Department already exists.")
    with tabs[1]:
        teachers=repo.rows("SELECT t.*,d.name home_department FROM teachers t LEFT JOIN departments d ON d.id=t.home_department_id ORDER BY t.full_name"); show_table(teachers)
        if user["role"] in ("System Administrator","Department Administrator"):
            deps=department_options()
            with st.form("new_teacher"):
                name=st.text_input("Teacher/person full name"); emp=st.text_input("Employee identifier (optional)"); des=st.text_input("Designation"); email=st.text_input("Institutional email (optional)"); dep=st.selectbox("Home department",[None]+deps,format_func=lambda d:"Not specified" if d is None else d["name"]); save=st.form_submit_button("Add teacher/person")
            if save:
                if not name.strip(): st.error("Name is required.")
                else:
                    try: tid=repo.execute("INSERT INTO teachers(full_name,employee_id,home_department_id,designation,email,created_at) VALUES(?,?,?,?,?,?)",(name.strip(),emp.strip() or None,dep["id"] if dep else None,des.strip() or None,email.strip() or None,now_iso())); log_ui(user,"CREATE","teacher",tid,after={"name":name,"department_id":dep["id"] if dep else None}); st.success("Teacher/person added."); st.rerun()
                    except Exception as e: st.error(str(e))
    with tabs[2]:
        courses=repo.rows("SELECT c.*,d.name teaching_department FROM courses c JOIN departments d ON d.id=c.teaching_department_id ORDER BY c.code"); show_table(courses)
        if user["role"] in ("System Administrator","Department Administrator","Course Coordinator"):
            deps=department_options()
            with st.form("new_course"):
                code=st.text_input("Course code"); title=st.text_input("Course title"); dep=st.selectbox("Course-owning / teaching department",deps,format_func=lambda d:d["name"]); typ=st.selectbox("Subject type",["Theory","Practical","Mixed"]); hours=st.text_input("Contact hours (required for Architecture sessional rule)"); save=st.form_submit_button("Add course")
            if save:
                if not code.strip() or not title.strip(): st.error("Course code and title are required.")
                else:
                    try: cid=repo.execute("INSERT INTO courses(code,title,teaching_department_id,subject_type,contact_hours) VALUES(?,?,?,?,?)",(code.strip(),title.strip(),dep["id"],typ,hours.strip() or None)); log_ui(user,"CREATE","course",cid,after={"code":code,"title":title,"teaching_department":dep["name"]}); st.success("Course added."); st.rerun()
                    except sqlite3.IntegrityError: st.error("Course code already exists in this teaching department.")
        sessions=visible_sessions(user)
        if sessions:
            st.markdown("**Attach a course to a class group and record its eligible registration count.**")
            sess=st.selectbox("Session",sessions,format_func=session_label,key="offer_sess"); groups=repo.rows("SELECT * FROM class_groups WHERE session_id=?",(sess["id"],)); courses=repo.rows("SELECT c.*,d.name teaching_department FROM courses c JOIN departments d ON d.id=c.teaching_department_id WHERE c.active=1 ORDER BY c.code")
            if groups and courses:
                with st.form("new_offering"):
                    group=st.selectbox("Class group",groups,format_func=lambda g:f"{g['academic_year']}-{g['semester']} {g['category']} (students: {g['student_count']})")
                    course=st.selectbox("Course",courses,format_func=lambda c:f"{c['code']} — {c['title']} ({c['teaching_department']})"); n=st.number_input("Course registration count",min_value=0,step=1,value=0); confirm=st.checkbox("Confirm count against course registration record"); rem=st.text_input("Remarks"); save=st.form_submit_button("Attach course")
                if save:
                    try: oid=repo.execute("INSERT INTO course_offerings(group_id,course_id,registration_count,count_confirmed,remarks) VALUES(?,?,?,?,?)",(group["id"],course["id"],int(n),int(confirm),rem.strip() or None)); log_ui(user,"CREATE","course_offering",oid,after={"group_id":group["id"],"course_id":course["id"],"registration_count":n}); st.success("Course offering created."); st.rerun()
                    except sqlite3.IntegrityError: st.error("This course is already attached to the selected group.")
            elif not groups: st.info("Create a class group first.")
            else: st.info("Add a course first.")

def page_rates(user):
    st.header("Financial rates and versions")
    st.warning("Rates seeded from the supplied prompt are drafts and inactive. Before financial use, confirm the exact official PDF wording, page/item, effective date, minimum/maximum scope, and applicability. Unapproved rates cannot generate payable amounts.")
    rates=repo.rows("SELECT r.*,s.title source_title FROM rates r LEFT JOIN sources s ON s.id=r.source_id ORDER BY r.item_name,r.version DESC"); show_table(rates)
    if user["role"] in ("System Administrator","Department Administrator"):
        with st.expander("Add a new rate version",expanded=False):
            with st.form("add_rate"):
                item=st.text_input("Normalized item name"); wording=st.text_input("Original Bengali / official wording"); amount=st.text_input("Rate amount (BDT)"); unit=st.text_input("Payment unit"); dist=st.selectbox("Distribution",["individual","equal sharing","board sharing"]); minimum=st.text_input("Minimum amount (optional)"); maxamt=st.text_input("Maximum amount (optional)"); mins=st.text_input("Minimum scope"); maxs=st.text_input("Maximum scope"); applicable=st.selectbox("Applicability",["REQUIRES AUTHORIZED VERIFICATION","Regular","Review","Regular and Review","Special"]); effective=st.date_input("Effective date",value=None); expiry=st.date_input("Expiry date (optional)",value=None); page=st.text_input("Official PDF page"); sourceitem=st.text_input("Official item number/reference"); reason=st.text_area("Change reason (required)"); save=st.form_submit_button("Create draft rate version")
            if save:
                try:
                    amt=money(amount)
                    if not item.strip() or not unit.strip() or amt<0 or not reason.strip(): raise ValueError("Item, unit, nonnegative rate and change reason are required.")
                    old=repo.one("SELECT MAX(version) v FROM rates WHERE item_name=?",(item.strip(),)); version=int(old["v"] or 0)+1
                    rid=repo.execute("INSERT INTO rates(item_name,source_wording,amount,unit,minimum_amount,maximum_amount,min_scope,max_scope,distribution,study_level,applicability,effective_date,expiry_date,source_page,source_item,approval_status,version,change_reason,created_by,created_at,active) VALUES(?,?,?,?,?,?,?,?,?,'Undergraduate',?,?,?,?,?,'Draft',?,?,?, ?,0)",(item.strip(),wording.strip() or None,str(amt),unit.strip(),str(money(minimum)) if minimum.strip() else None,str(money(maxamt)) if maxamt.strip() else None,mins.strip() or None,maxs.strip() or None,dist,applicable,effective.isoformat() if effective else None,expiry.isoformat() if expiry else None,page.strip() or None,sourceitem.strip() or None,version,reason.strip(),user["id"],now_iso()))
                    log_ui(user,"CREATE","rate",rid,after={"item":item,"amount":str(amt),"version":version},reason=reason); st.success(f"Draft rate version {version} created."); st.rerun()
                except Exception as e: st.error(str(e))
        pending=[r for r in rates if r["approval_status"]!="Approved"]
        if pending:
            st.markdown("**Review a rate**")
            sel=st.selectbox("Draft rate",pending,format_func=lambda r:f"#{r['id']} {r['item_name']} v{r['version']} — Tk. {r['amount']}")
            source_title=st.text_input("Official source document title",value="Updated new Exam bill rate.pdf",key="src_title"); source_ref=st.text_input("Circular/reference metadata (do not invent)",key="src_ref"); source_date=st.date_input("Official effective date (confirm from PDF)",value=None,key="src_date"); source_page=st.text_input("Confirmed official PDF page",value=sel["source_page"] or "",key="src_page"); source_item=st.text_input("Confirmed official item number/reference",value=sel["source_item"] or "",key="src_item"); source_remarks=st.text_input("Page/item confirmation notes",key="src_notes")
            c1,c2=st.columns(2)
            if c1.button("Approve and activate selected rate",key="approve_rate"):
                try:
                    require_role(user,["Authorized Reviewer/Approver","System Administrator"])
                    if not source_title.strip() or not source_remarks.strip() or not source_page.strip() or not source_item.strip(): raise ValueError("Enter confirmed official source page and item, and source confirmation notes before approval.")
                    sid=repo.execute("INSERT INTO sources(title,document_ref,effective_date,remarks,created_at) VALUES(?,?,?,?,?)",(source_title.strip(),source_ref.strip() or None,source_date.isoformat() if source_date else None,source_remarks.strip(),now_iso()))
                    repo.execute("UPDATE rates SET source_id=?,effective_date=COALESCE(effective_date,?),source_page=?,source_item=?,approval_status='Approved',approved_by=?,active=1,applicability=CASE WHEN applicability='REQUIRES AUTHORIZED VERIFICATION' THEN 'Approved' ELSE applicability END WHERE id=?",(sid,source_date.isoformat() if source_date else None,source_page.strip(),source_item.strip(),user["id"],sel["id"]))
                    log_ui(user,"APPROVE_ACTIVATE","rate",sel["id"],before={"approval_status":sel["approval_status"]},after={"approval_status":"Approved","source_id":sid}); st.success("Rate approved and activated."); st.rerun()
                except Exception as e: st.error(str(e))
            if c2.button("Deactivate selected rate",key="deactivate_rate"):
                try:
                    require_role(user,["System Administrator"]); repo.execute("UPDATE rates SET active=0 WHERE id=?",(sel["id"],)); log_ui(user,"DEACTIVATE","rate",sel["id"]); st.success("Rate deactivated."); st.rerun()
                except Exception as e: st.error(str(e))

def page_duties(user):
    st.header("Duty assignments and quantities")
    sessions=visible_sessions(user)
    if not sessions: st.info("Create or obtain access to an examination session first."); return
    sess=st.selectbox("Examination session",sessions,format_func=session_label,key="duty_sess")
    groups=repo.rows("SELECT * FROM class_groups WHERE session_id=?",(sess["id"],)); teachers=repo.rows("SELECT * FROM teachers WHERE active=1 ORDER BY full_name")
    # Show approved rates for normal use; if none are approved, expose drafts only for
    # building provisional assignments after the admin has explicitly reviewed them.
    rates=repo.rows("SELECT * FROM rates WHERE active=1 AND approval_status='Approved' AND study_level='Undergraduate' ORDER BY item_name")
    provisional_rates=False
    if not rates:
        rates=repo.rows("SELECT * FROM rates WHERE study_level='Undergraduate' AND approval_status='Draft' ORDER BY item_name")
        provisional_rates=bool(rates)
    duties=repo.rows("SELECT d.*,t.full_name teacher,r.amount rate_amount,r.version rate_version,r.approval_status FROM duty_assignments d JOIN teachers t ON t.id=d.teacher_id JOIN rates r ON r.id=d.rate_id WHERE d.session_id=? ORDER BY d.id DESC",(sess["id"],)); show_table(duties)
    if sess["status"] in ("Approved","Closed"): st.warning("This session is approved/closed; assignments are read-only."); return
    if user["role"] not in ("System Administrator","Department Administrator","Course Coordinator"): st.info("Your role does not permit duty assignment changes."); return
    if not rates: st.warning("No undergraduate rates are available. Ask the System Administrator to initialize or import the rate schedule."); return
    if provisional_rates:
        st.warning("PROVISIONAL MODE: no rates are approved yet. You can prepare draft duty assignments, but bill amounts remain blocked from payable totals until an authorized reviewer approves the rates.")
    offerings=[]
    if groups:
        ids=[g["id"] for g in groups]; offerings=repo.rows("SELECT o.*,c.code,c.title FROM course_offerings o JOIN courses c ON c.id=o.course_id WHERE o.group_id IN ("+",".join("?" for _ in ids)+") ORDER BY c.code",ids)
    with st.form("new_duty"):
        group=st.selectbox("Class group (optional)",[None]+groups,format_func=lambda g:"Not group-specific" if g is None else f"{g['academic_year']}-{g['semester']} {g['category']} (count {g['student_count']})")
        offering=st.selectbox("Course offering (optional)",[None]+offerings,format_func=lambda o:"Not course-specific" if o is None else f"{o['code']} — {o['title']} (registered {o['registration_count']})")
        teacher=st.selectbox("Duty-performing teacher/person",teachers,format_func=lambda t:t["full_name"]); rate=st.selectbox("Rate (approved rates first; drafts are provisional)",rates,format_func=lambda r:f"{r['item_name']} — Tk. {r['amount']} / {r['unit']} [{r['approval_status']}] ")
        q=st.text_input("Exact payable quantity (do not infer from class size)"); unit=st.text_input("Recorded quantity unit",value=rate["unit"]); scope=st.text_input("Calculation scope",value="Per recorded duty assignment"); formula=st.text_input("Formula / explanation",value="eligible quantity × approved rate"); assignees=st.multiselect("Eligible assignees for shared distribution",teachers,format_func=lambda t:t["full_name"]); rem=st.text_area("Supporting evidence / remarks"); save=st.form_submit_button("Create duty assignment")
    if save:
        try:
            qty=money(q)
            if qty<=0: raise ValueError("Quantity must be positive.")
            if not unit.strip() or not scope.strip() or not formula.strip(): raise ValueError("Unit, scope and formula are required.")
            if rate["distribution"]!="individual" and not assignees: raise ValueError("Select eligible assignees for a shared rate.")
            if rate["item_name"]=="Sessional (Architecture special rate)":
                if not offering: raise ValueError("Architecture sessional needs a course offering.")
                course=repo.one("SELECT c.*,d.name dept FROM courses c JOIN departments d ON d.id=c.teaching_department_id WHERE c.id=(SELECT course_id FROM course_offerings WHERE id=?)",(offering["id"],))
                if course["dept"]!="Architecture": raise ValueError("Architecture-specific rate cannot be applied to other departments.")
                if not course["contact_hours"]: raise ValueError("Architecture sessional calculation requires confirmed contact hours.")
            if rate["item_name"]=="Answer-script scrutiny":
                yr=group["academic_year"] if group else None; sem=group["semester"] if group else None
                if yr and sem:
                    current=repo.rows("SELECT DISTINCT teacher_id FROM duty_assignments WHERE session_id=? AND group_id IN (SELECT id FROM class_groups WHERE session_id=? AND academic_year=? AND semester=?) AND rate_id=?",(sess["id"],sess["id"],yr,sem,rate["id"]))
                    ids=[x["teacher_id"] for x in current]+[teacher["id"]]
                    issues=validate_scrutinizers(yr,sem,ids)
                    if issues: raise ValueError(" ".join(issues))
            did=repo.execute("INSERT INTO duty_assignments(session_id,group_id,offering_id,teacher_id,rate_id,item_name,quantity,unit,eligible_assignees,formula,scope,status,remarks,created_by,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,'Draft',?,?,?,?)",(sess["id"],group["id"] if group else None,offering["id"] if offering else None,teacher["id"],rate["id"],rate["item_name"],str(qty),unit.strip(),json.dumps([x["id"] for x in assignees] if assignees else [teacher["id"]]),formula.strip(),scope.strip(),rem.strip() or None,user["id"],now_iso(),now_iso()))
            log_ui(user,"CREATE","duty_assignment",did,after={"session_id":sess["id"],"teacher_id":teacher["id"],"rate_id":rate["id"],"quantity":str(qty)}); st.success("Duty assignment created."); st.rerun()
        except Exception as e: st.error(str(e))

def page_bills(user):
    st.header("Individual and consolidated bills")
    sessions=visible_sessions(user)
    if not sessions: st.info("No accessible sessions."); return
    sess=st.selectbox("Session",sessions,format_func=session_label,key="bill_sess")
    if not can_edit_session(user,sess): st.error("Unauthorized session."); return
    if user["role"] in ("System Administrator","Department Administrator","Course Coordinator"):
        teachers=repo.rows("SELECT DISTINCT t.* FROM teachers t JOIN duty_assignments d ON d.teacher_id=t.id WHERE d.session_id=? ORDER BY t.full_name",(sess["id"],))
        if teachers:
            teacher=st.selectbox("Generate individual bill for",teachers,format_func=lambda t:t["full_name"])
            if st.button("Calculate / refresh draft bill",key="calc_bill"):
                try:
                    result=generate_teacher_bill(sess["id"],teacher["id"]); log_ui(user,"CALCULATE","bill_header",result["id"],after=result); st.success(f"Draft bill {result['reference_no']} calculated: {bd_money(result['total'])}.")
                    if result["issues"]: st.warning("Not included in payable total: " + " | ".join(result["issues"]))
                    st.rerun()
                except Exception as e: st.error(str(e))
    bills=repo.rows("SELECT b.*,t.full_name teacher,s.exam_name FROM bill_headers b JOIN teachers t ON t.id=b.teacher_id JOIN exam_sessions s ON s.id=b.session_id WHERE b.session_id=? ORDER BY b.id DESC",(sess["id"],)); show_table(bills)
    if bills:
        b=st.selectbox("Bill detail",bills,format_func=lambda x:f"{x['reference_no']} — {x['teacher']} — {bd_money(x['total'])} [{x['status']}]")
        lines=repo.rows("SELECT l.*,d.group_id,d.offering_id FROM bill_lines l JOIN duty_assignments d ON d.id=l.duty_id WHERE l.bill_id=?",(b["id"],)); show_table(lines)
        if lines:
            df=pd.DataFrame(lines); st.download_button("Export selected teacher bill (.xlsx)",workbook_bytes(df,"Individual Bill",f"{APP_TITLE} — {b['reference_no']}"),file_name=f"teacher_bill_{b['id']}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    # consolidated sums are from the same persisted payable line snapshots
    consolidated=repo.rows("SELECT t.full_name teacher,d.item_name,d.quantity,d.unit,r.amount rate,r.version,l.gross_amount,l.adjustment,l.payable_amount,l.validation_status,b.reference_no,g.academic_year,g.semester,g.category,c.code course,c.title course_title,td.name teaching_department FROM bill_lines l JOIN bill_headers b ON b.id=l.bill_id JOIN teachers t ON t.id=b.teacher_id JOIN duty_assignments d ON d.id=l.duty_id JOIN rates r ON r.id=d.rate_id LEFT JOIN class_groups g ON g.id=d.group_id LEFT JOIN course_offerings o ON o.id=d.offering_id LEFT JOIN courses c ON c.id=o.course_id LEFT JOIN departments td ON td.id=c.teaching_department_id WHERE b.session_id=? ORDER BY t.full_name,d.id",(sess["id"],))
    st.subheader("Consolidated examination bill / audit register"); show_table(consolidated)
    if consolidated:
        total=sum((money(r["payable_amount"]) for r in consolidated),Decimal("0")); st.metric("Consolidated payable total (persisted bill lines)",bd_money(total))
        st.download_button("Export consolidated audit register (.xlsx)",workbook_bytes(pd.DataFrame(consolidated),"Consolidated Bill",f"{APP_TITLE} — Session {sess['id']}"),file_name=f"consolidated_bill_session_{sess['id']}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

def page_import(user):
    st.header("Excel import and reconciliation")
    st.warning("Import preview is read-only. This version records inspected source rows for reconciliation; it does not automatically map or commit unknown workbook layouts into financial records.")
    upload=st.file_uploader("Upload historical .xlsm or .xlsx workbook",type=["xlsm","xlsx"])
    if upload:
        try:
            sheets,notes=inspect_excel(upload); st.write(f"**File:** {upload.name} | **Sheets:** {', '.join(sheets)}")
            for note in notes: st.warning(note)
            sheet=st.selectbox("Preview sheet",list(sheets)); df=sheets[sheet]; st.dataframe(df.head(100),use_container_width=True)
            st.caption("Rows are displayed as imported; formulas/macros are not executed. Review and map columns before using values in bills.")
            if st.button("Create reconciliation batch from this workbook"):
                batch=repo.execute("INSERT INTO import_batches(filename,imported_by,status,summary,created_at) VALUES(?,?,'Previewed',?,?)",(upload.name,user["id"],json.dumps({"sheets":list(sheets),"notes":notes}),now_iso()))
                # Preserve source row text without guessing field meanings.
                with repo.connect() as c:
                    c.execute("BEGIN IMMEDIATE")
                    try:
                        for sh,frame in sheets.items():
                            for ix,row in frame.iterrows():
                                vals=[str(x) if pd.notna(x) else "" for x in row.tolist()]
                                if not any(vals): continue
                                c.execute("INSERT INTO reconciliation(batch_id,sheet_name,source_row,source_item,validation_note,status) VALUES(?,?,?,?,?,'Unresolved')",(batch,sh,str(ix+1),json.dumps(vals,ensure_ascii=False),"Source row retained for manual column mapping; rate/formula not inferred."))
                        c.execute("COMMIT")
                    except Exception: c.execute("ROLLBACK"); raise
                log_ui(user,"IMPORT_PREVIEW","import_batch",batch,after={"filename":upload.name,"sheets":list(sheets)}); st.success(f"Recorded preview batch #{batch}. Rows remain unresolved until mapped and reviewed."); st.rerun()
        except Exception as e: st.error(f"Could not inspect workbook safely: {e}")
    st.subheader("Import batches"); batches=repo.rows("SELECT * FROM import_batches ORDER BY id DESC"); show_table(batches)
    st.subheader("Reconciliation exceptions"); rows=repo.rows("SELECT * FROM reconciliation ORDER BY id DESC LIMIT 1000"); show_table(rows)
    if rows and user["role"] in ("Authorized Reviewer/Approver","System Administrator"):
        selected=st.selectbox("Review row",rows,format_func=lambda r:f"#{r['id']} {r['sheet_name']} row {r['source_row']} — {r['status']}")
        with st.form("recon_decision"):
            status=st.selectbox("Decision",["Confirmed match","Historical-rate difference","Formula difference","Quantity difference","Distribution difference","Rounding difference","Missing record","Accepted discrepancy","Unresolved"]); note=st.text_area("Reviewer decision / evidence",required=True); save=st.form_submit_button("Save reconciliation decision")
        if save:
            try:
                repo.execute("UPDATE reconciliation SET status=?,validation_note=?,reviewer_id=?,decision_date=? WHERE id=?",(status,note,user["id"],now_iso(),selected["id"])); log_ui(user,"RECONCILIATION_DECISION","reconciliation",selected["id"],before={"status":selected["status"]},after={"status":status,"note":note}); st.success("Decision saved."); st.rerun()
            except Exception as e: st.error(str(e))

def page_approvals(user):
    st.header("Review, approval and audit logs")
    st.subheader("Bills and workflow history")
    sessions=visible_sessions(user); ids=[s["id"] for s in sessions]
    if ids:
        ph=",".join("?" for _ in ids)
        show_table(repo.rows(f"SELECT a.*,u.email actor_email FROM approvals a JOIN users u ON u.id=a.actor_id WHERE a.session_id IN ({ph}) ORDER BY a.id DESC",ids))
        show_table(repo.rows(f"SELECT b.*,t.full_name teacher FROM bill_headers b JOIN teachers t ON t.id=b.teacher_id WHERE b.session_id IN ({ph}) ORDER BY b.id DESC",ids))
    if user["role"] in ("System Administrator","Authorized Reviewer/Approver"):
        st.subheader("Audit trail"); logs=repo.rows("SELECT id,actor_email,action,entity,record_id,before_json,after_json,reason,created_at FROM audit_logs ORDER BY id DESC LIMIT 1000"); show_table(logs)
    else: st.info("Audit logs are restricted to authorized reviewers and system administrators.")

def page_reports(user):
    st.header("Reports and exports")
    sessions=visible_sessions(user); ids=[s["id"] for s in sessions]
    if not ids: st.info("No accessible records."); return
    ph=",".join("?" for _ in ids)
    reports={"Teacher-wise summary":f"SELECT t.full_name teacher,b.status,SUM(CAST(b.total AS NUMERIC)) total,COUNT(*) bills FROM bill_headers b JOIN teachers t ON t.id=b.teacher_id WHERE b.session_id IN ({ph}) GROUP BY t.full_name,b.status", "Duty-wise summary":f"SELECT l.item_name,l.validation_status,SUM(CAST(l.payable_amount AS NUMERIC)) total,COUNT(*) lines FROM bill_lines l JOIN bill_headers b ON b.id=l.bill_id WHERE b.session_id IN ({ph}) GROUP BY l.item_name,l.validation_status", "Rate schedule": "SELECT item_name,source_wording,amount,currency,unit,minimum_amount,maximum_amount,distribution,applicability,approval_status,version,source_page,source_item,effective_date,active FROM rates ORDER BY item_name,version DESC", "Validation exceptions":f"SELECT l.*,b.reference_no FROM bill_lines l JOIN bill_headers b ON b.id=l.bill_id WHERE b.session_id IN ({ph}) AND l.validation_status!='VALID'"}
    report=st.selectbox("Report",list(reports)); rows=repo.rows(reports[report],ids if "?" in reports[report] else ())
    show_table(rows)
    if rows: st.download_button("Download report (.xlsx)",workbook_bytes(pd.DataFrame(rows),"Report",f"{APP_TITLE} — {report}"),file_name=re.sub(r"[^A-Za-z0-9_-]+","_",report.lower())+".xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

def page_users_settings(user):
    st.header("Users, roles and settings")
    if user["role"]!="System Administrator": st.error("System Administrator role required."); return
    users=repo.rows("SELECT id,email,full_name,role,department_id,verified,active,failed_logins,created_at FROM users ORDER BY id"); show_table(users)
    st.subheader("Change account status / role")
    if users:
        target=st.selectbox("Account",users,format_func=lambda u:f"{u['email']} — {u['role']} (verified={bool(u['verified'])}, active={bool(u['active'])})")
        roles=st.selectbox("Role",ROLES,index=ROLES.index(target["role"])); deps=department_options(); dep=st.selectbox("Department scope",[None]+deps,format_func=lambda d:"No department scope" if d is None else d["name"]); active=st.checkbox("Account active",value=bool(target["active"])); verified=st.checkbox("Email ownership verified (only after trusted verification evidence)",value=bool(target["verified"])); reason=st.text_input("Reason for account change")
        if st.button("Save account controls"):
            try:
                if target["id"]==user["id"] and roles!=user["role"]: raise ValueError("Do not change your own role in this session.")
                if not reason.strip(): raise ValueError("Reason required for account changes.")
                repo.execute("UPDATE users SET role=?,department_id=?,active=?,verified=? WHERE id=?",(roles,dep["id"] if dep else None,int(active),int(verified),target["id"])); log_ui(user,"UPDATE_ACCOUNT","user",target["id"],before=target,after={"role":roles,"department_id":dep["id"] if dep else None,"active":active,"verified":verified},reason=reason); st.success("Account controls updated."); st.rerun()
            except Exception as e: st.error(str(e))
    st.subheader("Runtime configuration status")
    for key in ["EXAM_BILL_DB","EXAM_BILL_ADMIN_BOOTSTRAP_TOKEN","SMTP_HOST","SMTP_FROM","SMTP_PASSWORD","SMTP_PORT","SMTP_USER"]:
        value=config_value(key); st.write(f"**{key}:** {'configured' if value else 'not configured'}")
    st.caption("Secrets are read from environment variables and are never displayed. Local SQLite data may not persist on some cloud hosts; configure EXAM_BILL_DB to a durable writable volume.")

def main():
    st.set_page_config(page_title=APP_TITLE,page_icon="📑",layout="wide",initial_sidebar_state="expanded")
    st.markdown("""<style>.block-container{padding-top:1.3rem;max-width:1500px}.stMetric{border:1px solid #d8dee8;padding:12px;border-radius:8px}.app-foot{color:#64748b;font-size:.82rem;border-top:1px solid #ddd;padding-top:10px}</style>""",unsafe_allow_html=True)
    if "user" not in st.session_state or not st.session_state.user:
        page_login(); st.markdown(f"<div class='app-foot'>{DEVELOPER}</div>",unsafe_allow_html=True); return
    user=st.session_state.user
    # Re-check current account state on each rerun, rather than trusting session state alone.
    current=repo.one("SELECT verified,active,role,department_id FROM users WHERE id=?",(user["id"],))
    if not current or not current["active"]:
        st.session_state.pop("user",None); st.error("Your account is inactive. Please sign in again."); st.rerun()
    user.update({k:current[k] for k in ("verified","active","role","department_id")})
    with st.sidebar:
        st.title("Exam Bill Calculator")
        st.caption(f"Signed in: {user['full_name']}")
        st.caption(f"Role: {user['role']}")
        st.caption("Scope: " + (repo.one("SELECT name FROM departments WHERE id=?",(user["department_id"],))["name"] if user.get("department_id") and repo.one("SELECT name FROM departments WHERE id=?",(user["department_id"],)) else "System-wide / unassigned"))
        pages=["Dashboard","Examination sessions","Departments and class groups","Courses and teachers","Duty assignments","Financial rates","Individual and consolidated bills","Excel import and reconciliation","Review and approval","Reports and exports","User administration","Settings","About"]
        page=st.radio("Navigation",pages,label_visibility="collapsed")
        if st.button("Log out",use_container_width=True): repo.audit(user,"LOGOUT","user",user["id"]); st.session_state.pop("user",None); st.rerun()
        st.divider(); st.caption(DEVELOPER)
    try:
        if page=="Dashboard": page_dashboard(user)
        elif page=="Examination sessions": page_sessions(user)
        elif page=="Departments and class groups": page_sessions(user)
        elif page=="Courses and teachers": page_departments_teachers_courses(user)
        elif page=="Duty assignments": page_duties(user)
        elif page=="Financial rates": page_rates(user)
        elif page=="Individual and consolidated bills": page_bills(user)
        elif page=="Excel import and reconciliation": page_import(user)
        elif page=="Review and approval": page_approvals(user)
        elif page=="Reports and exports": page_reports(user)
        elif page=="User administration": page_users_settings(user)
        elif page=="Settings": page_users_settings(user)
        elif page=="About":
            st.header(APP_TITLE); st.write("A traceable workflow for undergraduate examination bills, financial-rate versioning, reconciliation and audit.")
            st.markdown(f"**{DEVELOPER}**")
            st.warning("This application is not endorsed by DUET. Rates are initially seeded as unapproved drafts. Confirm the official revised remuneration order and all applicability interpretations before financial use.")
            st.write("Postgraduate and PhD rates are not active in this undergraduate module. Imported macros are never executed. Exports are drafts unless the application workflow records approval.")
    except PermissionError as e: st.error(str(e))
    except Exception as e:
        logger.exception("UI action failed")
        st.error(f"Operation could not be completed safely: {e}")
    st.markdown(f"<div class='app-foot'>{DEVELOPER} · Not an official DUET endorsement</div>",unsafe_allow_html=True)

if __name__ == "__main__":
    main()
