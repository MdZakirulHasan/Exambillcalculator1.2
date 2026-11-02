#!/usr/bin/env python3
"""DUET Exam Bill Calculator — streamlined final version.
Run: streamlit run exam_bill_calculator_1_2.py
Menu: Bill Rates | Teachers | Courses | Reports
"""
from __future__ import annotations
import os, io, sqlite3, zipfile, re
from datetime import datetime
from pathlib import Path
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

APP_TITLE = "DUET Exam Bill Calculator"
# Use a fresh schema filename. The previous deployment created tables with a different
# schema; reusing that SQLite file causes CREATE INDEX to fail on missing columns.
# Set EXAM_BILL_DB explicitly only if you have a database with this version's schema.
DB_PATH = os.getenv("EXAM_BILL_DB", "exam_bill_calculator_v3.sqlite3")
BILL_DEPARTMENTS = ['Civil Engineering (CE)', 'Electrical and Electronic Engineering (EEE)', 'Mechanical Engineering (ME)', 'Computer Science and Engineering (CSE)', 'Textile Engineering (TE)', 'Industrial & Production Engineering (IPE)', 'Architecture', 'Food Engineering (FE)', 'Chemical Engineering (ChE)', 'Materials and Metallurgical Engineering (MME)']
RATE_SEED = [('Question paper setting', 3600, 'per question paper per section', 'individual', 'Undergraduate'), ('Question-paper moderation', 2000, 'per moderator per year-semester; total fee and cap interpretation requires verification', 'equal sharing', 'Undergraduate'), ('Question-paper typing', 120, 'per question paper', 'equal sharing', 'Undergraduate'), ('Comparison, sketching, correction and related question-paper work', 1450, 'per question paper', 'equal sharing', 'Undergraduate'), ('Question-paper printing', 50, 'per question paper', 'equal sharing', 'Undergraduate'), ('Answer-script examination', 100, 'per answer script per section; minimum 1000 per examiner per section', 'individual', 'Undergraduate'), ('Assignment assessment', 50, 'per course per teacher per student', 'individual', 'Undergraduate'), ('Class test', 60, 'per class test per participating student', 'individual', 'Undergraduate'), ('Answer-script scrutiny', 20, 'per answer script per semester; minimum 1500 per scrutinizer per year-semester; max two scrutinizers', 'individual', 'Undergraduate'), ('Course-file preparation', 6000, 'per course file', 'equal sharing', 'Undergraduate'), ('Survey work (practical subjects)', 1000, 'per course per student', 'equal sharing', 'Undergraduate'), ('Sessional (practical subjects)', 320, 'per course per student; minimum 1600 per teacher per course', 'equal sharing', 'Undergraduate'), ('Sessional (Architecture special rate)', 115, 'per course per student per contact hour; minimum 1600 per teacher per course', 'equal sharing', 'Undergraduate'), ('Central viva (sessional/practical)', 150, 'per semester per student', 'board sharing', 'Undergraduate'), ('Industrial attachment/training', 750, 'per student', 'equal sharing', 'Undergraduate'), ('Thesis/project examination', 3000, 'per thesis/project per examiner', 'individual', 'Undergraduate'), ('Preliminary oral examination', 100, 'per examinee per examiner; year/semester and board scope requires confirmation', 'board sharing', 'Undergraduate'), ('Final oral examination', 250, 'per examinee per examiner; year/semester and board scope requires confirmation', 'board sharing', 'Undergraduate'), ('Thesis/project supervision', 6000, 'per thesis/project; maximum 24000 per supervisor', 'individual', 'Undergraduate'), ('Capstone project', 3000, 'per student', 'equal sharing', 'Undergraduate'), ('Chief invigilator', 1200, 'per hour', 'individual', 'Undergraduate'), ('Additional chief invigilator', 1100, 'per hour', 'individual', 'Undergraduate'), ('Other invigilator', 1000, 'per hour', 'individual', 'Undergraduate'), ('Controller of examinations', 1000, 'per hour', 'individual', 'Undergraduate'), ('Deputy controller of examinations', 900, 'per shift', 'individual', 'Undergraduate'), ('Programmer/assistant controller of examinations', 800, 'per shift', 'individual', 'Undergraduate'), ('Assistant officer to chief invigilator', 700, 'per shift', 'individual', 'Undergraduate'), ('Other officer', 550, 'per shift', 'individual', 'Undergraduate'), ('Assistant/third-class employee', 200, 'per hour', 'individual', 'Undergraduate'), ('Fourth-class employee', 175, 'per hour', 'individual', 'Undergraduate'), ('Grade-sheet preparation — theoretical', 60, 'per course per student', 'equal sharing', 'Undergraduate'), ('Grade-sheet preparation — practical', 50, 'per course per student', 'equal sharing', 'Undergraduate'), ('Grade-sheet scrutiny — theoretical', 30, 'per course per student', 'equal sharing', 'Undergraduate'), ('Grade-sheet scrutiny — practical', 12, 'per course per student', 'equal sharing', 'Undergraduate'), ('Tabulation', 90, 'per student', 'equal sharing', 'Undergraduate'), ('Advisory fee', 250, 'per semester per student', 'individual', 'Undergraduate'), ('Verification of final graduation result', 800, 'per student', 'individual', 'Undergraduate'), ('Computerized result preparation', 12, 'per course per student', 'equal sharing', 'Undergraduate'), ('Student-wise grade and GPA list verification', 24, 'per student', 'equal sharing', 'Undergraduate'), ('Chairman, examination committee honorarium', 4800, 'per semester; category eligibility requires confirmation', 'individual', 'Undergraduate'), ('Course coordinator', 3600, 'per year per semester; Review included with Regular; Special Examination exception', 'individual', 'Undergraduate')]
# Official department faculty pages checked on duet.ac.bd. Historical bill-only names are
# kept unassigned until their home department is verified; examination assignment is not home department.
OFFICIAL_TEACHERS = {
    "Food Engineering (FE)": [
        ("Dr. Mohammad Washim Dewan", "Head of Department"), ("Raju Ahmmed", "Assistant Professor (On Study Leave)"),
        ("Md. Junaeid Khan", "Assistant Professor (On Study Leave)"), ("Md. Abir Hossain", "Assistant Professor"),
        ("Md. Hasan Tarek Mondal", "Assistant Professor (On Study Leave)"), ("Dr. Md. Shofiul Azam", "Assistant Professor"),
        ("Lopa Aunsary", "Assistant Professor"), ("K. M. Mahdiuzzaman Sayed", "Lecturer"),
        ("Md. Zakirul Hasan", "Lecturer"), ("Md. Arman Ali", "Lecturer"),
    ],
    "Civil Engineering (CE)": [
        ("Dr. Md. Abu Taiyab", "Head of Department"), ("Dr. Mohammad Abdur Rashid", "Professor"),
        ("Dr. Md. Nazrul Islam", "Professor"), ("Dr. Ganesh Chandra Saha", "Professor"),
        ("Dr. Md. Khasro Miah", "Professor"), ("Dr. Md. Mokhlesur Rahman", "Professor"),
        ("Dr. Md. Kamal Hossain", "Professor"), ("Dr. Mohammad Nazim Uddin", "Professor"),
        ("Dr. Md. Akramul Alam", "Professor"), ("Dr. Mohammed Alauddin", "Professor"),
        ("Dr. Md. Abdus Salam", "Professor"), ("Dr. Mohammad Abdul Kader", "Professor"),
        ("Dr. Md. Monowar Hossain", "Professor"), ("Dr. Md. Shibly Anwar", "Professor"),
        ("Dr. Md. Kobir Hossain", "Associate Professor"), ("Dr. Md. Suman Mia", "Associate Professor (On Study Leave)"),
        ("Rokshana Pervin", "Assistant Professor"), ("Mr. Md. Rain Man Raja", "Assistant Professor (On Study Leave)"),
        ("Mr. Md. Rakibul Islam", "Assistant Professor (On Study Leave)"), ("Mr. Pronab Kumar Biswas", "Assistant Professor (On Study Leave)"),
        ("Dr. Md. Al-Amin Sikder", "Assistant Professor"), ("Mr. Shahab Uddin", "Assistant Professor (On Study Leave)"),
        ("Mr. Md. Ariful Islam", "Assistant Professor"), ("Mr. M. Afzal Hossain", "Assistant Professor"),
        ("Mr. Kazi Abu Manjur", "Assistant Professor"), ("Mr. Mahadi Hasan", "Assistant Professor"),
        ("Mr. Mahiman Zinnurain", "Assistant Professor"), ("Mr. Md. Nur Rahman", "Assistant Professor (On Study Leave)"),
        ("Ms. Tasnim Nishat Aishee", "Assistant Professor"), ("Mr. Md. Monjurul Hasan", "Lecturer (On Study Leave)"),
        ("Mr. Md Tarek Hasan", "Lecturer (On Study Leave)"), ("Monir Hossain", "Lecturer"),
    ],
    "Electrical and Electronic Engineering (EEE)": [
        ("Dr. Mohammad Jakir Hossain", "Head of Department"), ("Dr. Md. Anwarul Abedin", "Professor"),
        ("Dr. Md. Raju Ahmed", "Professor"), ("Dr. Md. Sharafat Hossain", "Professor"),
        ("Dr. Ruma", "Professor"), ("Dr. Md. Monirul Kabir", "Professor"),
    ],
    "Computer Science and Engineering (CSE)": [
        ("Dr. Momotaz Begum", "Head of Department"), ("Dr. Mohammod Abul Kashem", "Professor"),
        ("Dr. Md. Nasim Akhtar", "Professor"),
    ],
    "Mechanical Engineering (ME)": [("Dr. Md. Arefin Kowser", "Head of Department")],
    "Textile Engineering (TE)": [
        ("Dr. Md. Abdul Hannan", "Head of Department"), ("Dr. Abdus Shahid", "Professor"),
        ("Dr. Shahjalal Khandaker", "Professor"), ("Dr. Mohammad Zakaria", "Professor"),
        ("Dr. Muhammad Abdur Rahman Bhuiyan", "Professor"), ("Dr. Forkan Sarker", "Professor"),
        ("Dr. Muhammad Abdur Rashid", "Professor"), ("Dr. Shafiqul Islam", "Professor"),
        ("Dr. Anamul Hoque Bhuiyan", "Professor"), ("Dr. Md. Delwar Hossain", "Associate Professor"),
        ("Tanushree Saha", "Assistant Professor (On Study Leave)"), ("Md. Sumon Miah", "Assistant Professor (On Study Leave)"),
        ("Dr. Shamsuzzaman Sheikh", "Assistant Professor"), ("Abu Yousuf Mohammad Anwarul Azim", "Assistant Professor (On Study Leave)"),
        ("Ayub Ali", "Assistant Professor (On Study Leave)"), ("Md. Ashadujjaman", "Assistant Professor (On Study Leave)"),
        ("Md. Ariful Islam", "Assistant Professor"), ("Md. Mehedi Hasan", "Assistant Professor"),
        ("Md. Shablu Khan", "Assistant Professor"), ("Md. Rakibul Hasan", "Lecturer"), ("Md. Ahasan Habib", "Lecturer"),
    ],
    "Industrial & Production Engineering (IPE)": [
        ("Dr. Mohammed Moinul Islam", "Head of Department"), ("Prithbey Raj Dey", "Assistant Professor (On Study Leave)"),
        ("Mohammad Shahadath Hossain", "Assistant Professor"), ("Hasina Tabassum Chowdhury", "Assistant Professor (On Study Leave)"),
        ("Nilufa Yeasmin", "Assistant Professor"), ("Md. Shovon Zahid", "Assistant Professor (On Study Leave)"),
        ("Mahmud Khan", "Assistant Professor (On Study Leave)"), ("Chowdhury Sakib-Uz-Zaman", "Assistant Professor (On Study Leave)"),
        ("Shaim Mahamud", "Assistant Professor"), ("Md. Imran Hossain", "Assistant Professor (On Study Leave)"),
        ("Md. Ashfaqur Arefin", "Assistant Professor (On Study Leave)"), ("Dr. Mohammad Rashidul Hassan", "Assistant Professor"),
        ("Rotan Kumar Saha", "Assistant Professor"), ("Md. Touhiduzzaman", "Assistant Professor"),
        ("Faisal Mahmud", "Lecturer (On Study Leave)"), ("Mushaer Ahmed", "Lecturer"), ("Nur Md. Alif Ul Islam", "Lecturer"),
    ],
    "Architecture": [
        ("Dr. Bayezid Ismail Chowdhury", "Head of Department"), ("Joarder Hafiz Ullah", "Assistant Professor"),
        ("Dilruba Yasmin", "Assistant Professor (On Study Leave)"), ("Istiaque Ahmed", "Assistant Professor (On Study Leave)"),
        ("Farhana Ahsan", "Assistant Professor"),
    ],
    "Chemical Engineering (ChE)": [
        ("Dr. Md. Mostaqur Rahman", "Head of Department"), ("Mahmudul Hasan Khan Neon", "Assistant Professor (On Study Leave)"),
        ("Akib Atique Khan Dhrubo", "Assistant Professor"), ("Md. Shihabuzzaman Apon", "Assistant Professor (On Study Leave)"),
        ("Md. Jahirul Islam", "Assistant Professor (On Study Leave)"), ("S. M. Akimus Salehin", "Lecturer"),
        ("Nuruzzaman", "Lecturer"), ("Md. Rashedul Islam", "Lecturer"),
    ],
    "Materials and Metallurgical Engineering (MME)": [
        ("Dr. Hasan Mohammad Mostofa Afroz", "Head of Department"), ("Tasratur Reaj Neha", "Assistant Professor"),
        ("Md. Irfan Khan", "Assistant Professor"), ("Md. Raduanul Islam Mridul", "Assistant Professor"),
        ("Md. Jahidul Islam", "Assistant Professor"), ("Muhtasim Al Muyeed Jim", "Lecturer"),
        ("Md. Rakibul Hasan", "Lecturer"), ("Md. Shofiqul Islam", "Lecturer"), ("Md. Jewel Rana", "Lecturer"),
    ],
    "Chemistry": [
        ("Dr. Nour Md. Mofiz Uddin Khan", "Head of Department"), ("Dr. Zinia Nasreen", "Professor"),
        ("Dr. Ummey Rayhan", "Professor"), ("Dr. Md. Shahabuddin", "Associate Professor"),
        ("Md. Ikram Ul Hoque", "Assistant Professor (On Study Leave)"), ("Dr. Md. Sofiul Alom", "Assistant Professor"),
        ("Romana Nasrin", "Assistant Professor (On Study Leave)"), ("Md. Arif Hossain", "Assistant Professor"),
        ("Dr. Bikash Dev Nath", "Assistant Professor (On Study Leave)"), ("Dr. Md. Rafikul Islam", "Assistant Professor"),
    ],
    "Mathematics": [
        ("Dr. Most. Nasrin Akhter", "Head of Department"), ("Dr. Md. Abu Naim Sheikh", "Professor"),
        ("Dr. Md. Azmal Hossain", "Professor"), ("Dr. Md. Shirazul Hoque Mollah", "Professor"),
        ("Dr. Main Uddin Ahammad", "Professor"),
    ],
    "Physics": [
        ("Dr. Abu Talib Md. Kaosar Jamil", "Head of Department"), ("Dr. Syed Jamal Ahmed", "Professor"),
        ("Dr. Md. Kamal-Al-Hassan", "Professor"), ("Dr. Md. Sahab Uddin", "Professor"),
        ("Dr. Farah Deeba", "Professor"), ("Md. Rezaul Karim", "Associate Professor"),
        ("Mst. Afsana Azam", "Assistant Professor (On Study Leave)"), ("Md. Mamun-Or-Rashid", "Assistant Professor (On Study Leave)"),
        ("Abdulla Al Noman", "Assistant Professor"), ("Susmita Bhattacharjee", "Assistant Professor"),
        ("Prianka Mondol", "Lecturer (On Study Leave)"), ("Md. Saidur Rahman", "Lecturer"),
    ],
    "Humanities & Social Sciences": [
        ("Dr. Sanzida Rahman", "Head of Department"), ("Dr. A K M Zakaria", "Professor"),
        ("Dr. Md. Abdur Rahman Forhad", "Professor"), ("Dr. Most. Tahura Pervin", "Professor"),
        ("Dr. Fatema Sultana", "Professor"), ("Dr. Mohammad Faizul Haque", "Associate Professor"),
        ("Md. Mazharul Alam", "Assistant Professor"), ("Dr. Mir Md. Fazle Rabbi", "Assistant Professor"),
        ("Mamunur Rashid", "Assistant Professor"), ("Motiur Rahman", "Assistant Professor"),
        ("Afruza Haque", "Assistant Professor"), ("Nur Mohammad Khan", "Assistant Professor"),
        ("Md. Sawgat Khan", "Assistant Professor"), ("Sheikh Md. Rokonul Islam", "Assistant Professor"),
        ("Mohammad Kawsar Ahammed", "Lecturer"), ("Md. Hasan Al Azad", "Lecturer"),
    ],
}
# Minimum/maximum rules transcribed from the supplied Master Prompt's official-rate section.
RATE_LIMITS = {
    "Question-paper moderation": (2000, 5000),
    "Answer-script examination": (1000, 0),
    "Answer-script scrutiny": (1500, 0),
    "Sessional (practical subjects)": (1600, 0),
    "Sessional (Architecture special rate)": (1600, 0),
    "Thesis/project supervision": (0, 24000),
}
TEACHER_SEED = ['Anwar Parvez', 'Dr. Abdullah Iqbal', 'Dr. Farah Deeba', 'Dr. Main Uddin Ahammad', 'Dr. Md. Rafikul Islam', 'Dr. Md. Shahabuddin', 'Dr. Md. Shofiul Azam', 'Dr. Md. Sofiul Alom', 'Dr. Mohammad Washim Dewan', 'Dr. Sanzida Rahman', 'Dr.Khurshida Sharmin', 'K. M. Mahdiuzzaman Sayed', 'Lopa Aunsary', 'Md. Abir Hossain', 'Md. Arman Ali', 'Md. Nurunnabi Sohel', 'Md. Saidur Rahman', 'Md. Zakirul Hasan', 'Mr. Md. Mahbubur Rahman', 'Prianka Mondol', 'Susmita Bhattacharjee', 'Dr. Md. Rashed Mia', 'Md. Rezaul Karim', 'Mr. Saddam Hossen', 'Prof. Dr. Mohammad Washim Dewan', 'Prof. Dr. Mohammad Zoynal Abedin', 'Dr. Md. Anowar Hossain', 'Dr. Md. Mozammel Hoque', 'Md. Mahbubur Rahman', 'Md. Mahfujur Rahman', 'Md. Zihan Babu', 'Mr. Motiur Rahman', 'Sheikh Md. Rokonul Islam', 'Sumita Ghosh', 'Mr. Md. Shariful Alam Emon', 'Prof. Dr. Hasan Md. Mostofa Afroz', 'Prof. Dr. Md. Arefin Kowser', 'Prof. Dr. Md. Shirazul Hoque Mollah', 'Saifullah Mahmud', 'Dr. Mohammad Jakir Hossain', 'Mamunur Rashid', 'Md. Mohin Uddin', 'Md. Simul Hosan Talukder', 'Ms. Afruza Haque', 'Prof. Dr. Abdullah Iqbal']
COURSE_SEED = ['FE 2002', 'FE 2004', 'FE 2101', 'FE 2102', 'FE 2501', 'FE 2502', 'FE 2802', 'FE 3501', 'FE 3502', 'FE 3601', 'FE 3602', 'FE 3603', 'FE 3604', 'FE 3605', 'FE 4501', 'FE 4504', 'FE 4601', 'FE 4602', 'FE 4604', 'FE 4605', 'FE 4606', 'FE 4615', 'FE 4617', 'HSS 2801', 'ME 2801', 'ME 3606', 'ME 3801', 'ME 3802']
TEACHING_DEPARTMENTS = ['Civil Engineering (CE)', 'Electrical and Electronic Engineering (EEE)', 'Mechanical Engineering (ME)', 'Computer Science and Engineering (CSE)', 'Textile Engineering (TE)', 'Industrial & Production Engineering (IPE)', 'Architecture', 'Food Engineering (FE)', 'Chemical Engineering (ChE)', 'Materials and Metallurgical Engineering (MME)', 'Chemistry', 'Mathematics', 'Physics', 'Humanities & Social Sciences']


# ---------- Database ----------
def db():
    Path(DB_PATH).expanduser().resolve().parent.mkdir(parents=True, exist_ok=True)
    c=sqlite3.connect(DB_PATH, timeout=30)
    c.row_factory=sqlite3.Row
    c.execute("PRAGMA foreign_keys=ON")
    return c

def init_db():
    with db() as c:
        c.executescript("""
        CREATE TABLE IF NOT EXISTS departments(id INTEGER PRIMARY KEY, name TEXT UNIQUE NOT NULL);
        CREATE TABLE IF NOT EXISTS teachers(id INTEGER PRIMARY KEY, name TEXT NOT NULL, department_id INTEGER, designation TEXT DEFAULT '', active INTEGER DEFAULT 1, FOREIGN KEY(department_id) REFERENCES departments(id));
        CREATE TABLE IF NOT EXISTS courses(id INTEGER PRIMARY KEY, code TEXT NOT NULL, title TEXT NOT NULL, department_id INTEGER NOT NULL, active INTEGER DEFAULT 1, FOREIGN KEY(department_id) REFERENCES departments(id));
        CREATE TABLE IF NOT EXISTS rates(id INTEGER PRIMARY KEY, item TEXT UNIQUE NOT NULL, rate REAL NOT NULL DEFAULT 0, unit TEXT NOT NULL, minimum REAL NOT NULL DEFAULT 0, maximum REAL NOT NULL DEFAULT 0, active INTEGER DEFAULT 1);
        CREATE TABLE IF NOT EXISTS bills(id INTEGER PRIMARY KEY, exam_key TEXT NOT NULL, department_id INTEGER NOT NULL, teacher_id INTEGER NOT NULL, degree TEXT NOT NULL, year_semester TEXT NOT NULL, exam_name TEXT NOT NULL, exam_year INTEGER NOT NULL, exam_type TEXT NOT NULL, created_at TEXT NOT NULL, FOREIGN KEY(department_id) REFERENCES departments(id), FOREIGN KEY(teacher_id) REFERENCES teachers(id));
        CREATE TABLE IF NOT EXISTS bill_lines(id INTEGER PRIMARY KEY, bill_id INTEGER NOT NULL, rate_id INTEGER NOT NULL, course_id INTEGER, description TEXT NOT NULL, quantity REAL NOT NULL, rate REAL NOT NULL, amount REAL NOT NULL, scope_key TEXT NOT NULL DEFAULT '', FOREIGN KEY(bill_id) REFERENCES bills(id) ON DELETE CASCADE, FOREIGN KEY(rate_id) REFERENCES rates(id), FOREIGN KEY(course_id) REFERENCES courses(id));
        CREATE INDEX IF NOT EXISTS ix_teacher_dept ON teachers(department_id,active);
        CREATE INDEX IF NOT EXISTS ix_course_dept ON courses(department_id,active);
        """)
        # Safe migration for databases created by earlier versions.
        line_cols={r[1] for r in c.execute("PRAGMA table_info(bill_lines)")}
        if "scope_key" not in line_cols:
            c.execute("ALTER TABLE bill_lines ADD COLUMN scope_key TEXT NOT NULL DEFAULT ''")
        for d in TEACHING_DEPARTMENTS: c.execute("INSERT OR IGNORE INTO departments(name) VALUES(?)",(d,))
        # Seed rates once. Amounts are starter values from the supplied app source and must be checked against current approved DUET order.
        for item, amount, unit, distribution, level in RATE_SEED:
            c.execute("INSERT OR IGNORE INTO rates(item,rate,unit,minimum,maximum,active) VALUES(?,?,?,?,?,1)",(item,float(amount),unit,0,0))
        # Seed official faculty entries with home departments; keep historical examiner names
        # unassigned unless their home department can be verified from the official source.
        dep_ids={r['name']:r['id'] for r in c.execute("SELECT id,name FROM departments")}
        for dep, faculty in OFFICIAL_TEACHERS.items():
            for name, designation in faculty:
                existing=c.execute("SELECT id FROM teachers WHERE lower(name)=lower(?) ORDER BY id LIMIT 1", (name,)).fetchone()
                if existing:
                    c.execute("UPDATE teachers SET department_id=?, designation=CASE WHEN designation='' THEN ? ELSE designation END WHERE id=?", (dep_ids[dep], designation, existing[0]))
                else:
                    c.execute("INSERT INTO teachers(name,department_id,designation) VALUES(?,?,?)", (name, dep_ids[dep], designation))
        # Do not seed the old bill's examiner names as unassigned teachers. They are not
        # reliable evidence of a person's home department; only verified departmental
        # faculty entries are preloaded. Users can add visiting/external examiners manually.
        for item, minimum, maximum in [(k,v[0],v[1]) for k,v in RATE_LIMITS.items()]:
            c.execute("UPDATE rates SET minimum=?,maximum=? WHERE item=? AND minimum=0 AND maximum=0", (minimum, maximum, item))
        # Course codes are seeded, with titles explicitly marked for confirmation. Prefix suggests awarding department only as an editable starting point.
        depts={r['name']:r['id'] for r in c.execute("SELECT id,name FROM departments")}
        for code in COURSE_SEED:
            prefix=code.split()[0]
            dn={"FE":"Food Engineering (FE)","ME":"Mechanical Engineering (ME)","HSS":"Humanities & Social Sciences"}.get(prefix,"Food Engineering (FE)")
            if not c.execute("SELECT 1 FROM courses WHERE code=? AND title=? AND department_id=?",(code,"TITLE TO BE CONFIRMED FROM APPROVED CURRICULUM",depts[dn])).fetchone():
                c.execute("INSERT INTO courses(code,title,department_id) VALUES(?,?,?)",(code,"TITLE TO BE CONFIRMED FROM APPROVED CURRICULUM",depts[dn]))

init_db()

def rows(sql, args=()):
    with db() as c: return [dict(x) for x in c.execute(sql,args).fetchall()]
def run(sql,args=()):
    with db() as c: cur=c.execute(sql,args); c.commit(); return cur.lastrowid

from decimal import Decimal, ROUND_HALF_UP

def money(x):
    return float(Decimal(str(x or 0)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))

def calc_amount(qty, rate, minimum, maximum):
    """Calculate a single duty line. For scope-sensitive items, enter one line per defined scope.
    0 minimum/maximum means no limit. Amount is gross quantity × rate, bounded by limits.
    """
    q=Decimal(str(qty or 0)); r=Decimal(str(rate or 0))
    if q < 0 or r < 0: raise ValueError("Quantity and rate cannot be negative.")
    value=q*r
    mn=Decimal(str(minimum or 0)); mx=Decimal(str(maximum or 0))
    if value > 0 and mn > 0: value=max(value,mn)
    if value > 0 and mx > 0: value=min(value,mx)
    return money(value)

def recalculate_scope(bill_id, rate_id, scope_key, minimum, maximum):
    """Apply minimum/maximum once to all lines sharing a declared scope; allocate proportionally."""
    with db() as c:
        found=c.execute("SELECT id,quantity,rate FROM bill_lines WHERE bill_id=? AND rate_id=? AND scope_key=? ORDER BY id",(bill_id,rate_id,scope_key)).fetchall()
        if not found: return
        gross=[Decimal(str(x[1]))*Decimal(str(x[2])) for x in found]
        total=sum(gross, Decimal("0"))
        if total <= 0: amounts=[Decimal("0") for _ in gross]
        else:
            bounded=total
            if Decimal(str(minimum or 0))>0: bounded=max(bounded,Decimal(str(minimum)))
            if Decimal(str(maximum or 0))>0: bounded=min(bounded,Decimal(str(maximum)))
            amounts=[]; running=Decimal("0")
            for i,g in enumerate(gross):
                if i==len(gross)-1: a=bounded-running
                else: a=(bounded*g/total).quantize(Decimal("0.01"),rounding=ROUND_HALF_UP); running+=a
                amounts.append(a)
        for line,amount in zip(found,amounts): c.execute("UPDATE bill_lines SET amount=? WHERE id=?",(float(amount),line[0]))
        c.commit()

def make_teacher_xlsx(bill, lines):
    wb=Workbook(); ws=wb.active; ws.title="Exam Bill"
    ws.sheet_view.showGridLines=False
    ws.merge_cells("A1:H1"); ws["A1"]="Dhaka University of Engineering & Technology, Gazipur"
    ws.merge_cells("A2:H2"); ws["A2"]="Gazipur - 1707"
    ws.merge_cells("A3:H3"); ws["A3"]="(Examination Related Remuneration)"
    ws.merge_cells("A5:H5"); ws["A5"]="Details of Examination Related Works"
    ws["A7"]="Degree"; ws["B7"]=bill["degree"]; ws["C7"]="Year/Semester"; ws["D7"]=bill["year_semester"]; ws["E7"]="Examination"; ws["F7"]=f"{bill['exam_name']} - {bill['exam_year']} ({bill['exam_type']})"
    ws["A8"]="Name"; ws["B8"]=bill["teacher_name"]; ws["C8"]="Designation"; ws["D8"]=bill["designation"] or ""; ws["E8"]="Department"; ws["F8"]=bill["department"]
    headers=["Sl. No.","Description of Work","Course Code","Course Title","Calculation Scope","Quantity","Rate (BDT)","Amount (BDT)"]
    for j,h in enumerate(headers,1): ws.cell(10,j,h)
    for i,line in enumerate(lines,1):
        r=10+i
        vals=[i,line['description'],line.get('course_code') or '',line.get('course_title') or '',line.get('scope_key') or '',line['quantity'],line['rate'],float(line['amount'])]
        for j,v in enumerate(vals,1): ws.cell(r,j,v)
    end=10+len(lines)+1
    ws.cell(end,7,"Total (BDT)"); ws.cell(end,8,f"=SUM(H11:H{end-1})" if lines else "=0")
    thin=Side(style="thin",color="777777")
    for row in ws.iter_rows(min_row=1,max_row=end,min_col=1,max_col=8):
        for cell in row:
            cell.alignment=Alignment(vertical="center",wrap_text=True)
            if cell.row in (1,2,3,5): cell.alignment=Alignment(horizontal="center",vertical="center")
            if cell.row in (10,end):
                cell.font=Font(bold=True)
                cell.fill=PatternFill("solid",fgColor="D9EAF7")
            if cell.row>=10: cell.border=Border(left=thin,right=thin,top=thin,bottom=thin)
    for r in (1,2,3,5): ws.cell(r,1).font=Font(bold=True,size=14 if r==1 else 11)
    ws.row_dimensions[1].height=25
    widths=[10,35,17,38,28,14,16,18]
    for i,w in enumerate(widths,1): ws.column_dimensions[get_column_letter(i)].width=w
    ws.freeze_panes="A11"
    out=io.BytesIO(); wb.save(out); return out.getvalue()

st.set_page_config(page_title=APP_TITLE,page_icon="📑",layout="wide")
st.title(APP_TITLE)
st.caption("DUET · Semester-wise audit input · Teacher-wise downloadable Excel bills")
menu=st.sidebar.radio("Menu",["Bill Rates","Teachers","Courses","Reports"])

if menu=="Bill Rates":
    st.header("Bill Rates")
    st.caption("Edit the rate, unit, minimum and maximum. Minimum/maximum = 0 means no limit. Verify all rates against the current approved DUET order before official use.")
    rate_rows=rows("SELECT id,item,rate,unit,minimum,maximum,active FROM rates ORDER BY id")
    df=pd.DataFrame(rate_rows)
    if not df.empty:
        # Keep the rate table limited to the five fields needed for financial review.
        display=df[["item","rate","unit","minimum","maximum"]].copy()
        edited=st.data_editor(display,hide_index=True,use_container_width=True,num_rows="fixed",column_config={
            "item":st.column_config.TextColumn("Bill Item",disabled=True),"rate":st.column_config.NumberColumn("Rate (Tk.)",min_value=0,step=1,format="%.2f"),"unit":st.column_config.TextColumn("Unit"),"minimum":st.column_config.NumberColumn("Minimum (Tk.)",min_value=0,step=100),"maximum":st.column_config.NumberColumn("Maximum (Tk.)",min_value=0,step=100)},key="rates_editor")
        if st.button("Save Bill Rates",type="primary"):
            try:
                with db() as c:
                    for i,r in edited.iterrows():
                        if float(r['maximum'] or 0)>0 and float(r['minimum'] or 0)>float(r['maximum']): raise ValueError(f"Minimum cannot exceed maximum: {r['item']}")
                        c.execute("UPDATE rates SET rate=?,unit=?,minimum=?,maximum=? WHERE id=?",(float(r['rate']),str(r['unit']).strip(),float(r['minimum']),float(r['maximum']),int(df.iloc[i]['id'])))
                    c.commit()
                st.success("Bill rates saved. New report lines use these saved rates."); st.rerun()
            except Exception as e: st.error(str(e))

elif menu=="Teachers":
    st.header("Teachers")
    deps=rows("SELECT id,name FROM departments ORDER BY name")
    depmap={d['name']:d['id'] for d in deps}
    # Department submenus: All Departments and each home department. No unassigned bucket.
    depname=st.selectbox("Department", ["All Departments"]+[d['name'] for d in deps], key="teacher_department_submenu")
    if depname=="All Departments":
        teachers=rows("SELECT t.id,t.name,t.designation,d.name AS department FROM teachers t JOIN departments d ON d.id=t.department_id WHERE t.active=1 ORDER BY d.name,t.name")
    else:
        teachers=rows("SELECT t.id,t.name,t.designation,d.name AS department FROM teachers t JOIN departments d ON d.id=t.department_id WHERE t.active=1 AND t.department_id=? ORDER BY t.name",(depmap[depname],))
    if teachers:
        st.dataframe(pd.DataFrame(teachers)[["department","name","designation"]],use_container_width=True,hide_index=True)
    else:
        st.info("No teachers are currently listed for this department. Add verified faculty below.")
    with st.expander("Add teacher to a department"):
        with st.form("add_teacher"):
            name=st.text_input("Teacher's full name",key="new_teacher_name")
            designation=st.text_input("Designation")
            dep=st.selectbox("Home department",[d['name'] for d in deps],index=([d['name'] for d in deps].index(depname) if depname in depmap else 0),key="new_teacher_dep")
            if st.form_submit_button("Add teacher"):
                if not name.strip(): st.error("Teacher name is required.")
                elif rows("SELECT id FROM teachers WHERE lower(name)=lower(?) AND department_id=? AND active=1",(name.strip(),depmap[dep])): st.error("This teacher is already listed in this department.")
                else:
                    run("INSERT INTO teachers(name,designation,department_id) VALUES(?,?,?)",(name.strip(),designation.strip(),depmap[dep]))
                    st.success(f"{name.strip()} added under {dep}."); st.rerun()
    with st.expander("Edit / deactivate teacher"):
        all_t=rows("SELECT t.*,d.name department FROM teachers t JOIN departments d ON d.id=t.department_id WHERE t.active=1 ORDER BY d.name,t.name")
        if all_t:
            chosen=st.selectbox("Teacher",all_t,format_func=lambda x:f"{x['department']} — {x['name']}")
            with st.form("edit_teacher"):
                newname=st.text_input("Full name",chosen['name'])
                newdes=st.text_input("Designation",chosen['designation'] or "")
                newdep=st.selectbox("Department",[d['name'] for d in deps],index=[d['name'] for d in deps].index(chosen['department']))
                deactivate=st.checkbox("Deactivate teacher")
                if st.form_submit_button("Save teacher"):
                    run("UPDATE teachers SET name=?,designation=?,department_id=?,active=? WHERE id=?",(newname.strip(),newdes.strip(),depmap[newdep],0 if deactivate else 1,chosen['id']))
                    st.success("Teacher updated."); st.rerun()

elif menu=="Courses":
    st.header("Courses")
    deps=rows("SELECT id,name FROM departments ORDER BY name"); depmap={d['name']:d['id'] for d in deps}
    depname=st.selectbox("Degree-awarding department",[d['name'] for d in deps],key="course_dept_filter")
    courses=rows("SELECT c.*,d.name department FROM courses c JOIN departments d ON d.id=c.department_id WHERE c.active=1 AND c.department_id=? ORDER BY c.code,c.title",(depmap[depname],))
    st.dataframe(pd.DataFrame(courses),use_container_width=True,hide_index=True)
    with st.expander("Add course / syllabus version"):
        with st.form("add_course"):
            code=st.text_input("Course code (required)"); title=st.text_input("Course title (required)"); course_dep=st.selectbox("Degree-awarding department",[d['name'] for d in deps],index=[d['name'] for d in deps].index(depname));
            if st.form_submit_button("Add course"):
                if not code.strip() or not title.strip(): st.error("Both course code and course title are mandatory.")
                else:
                    run("INSERT INTO courses(code,title,department_id) VALUES(?,?,?)",(code.strip().upper(),title.strip(),depmap[course_dep])); st.success("Course added. Same code with another title is allowed."); st.rerun()
    with st.expander("Edit / deactivate course"):
        all_c=rows("SELECT c.*,d.name department FROM courses c JOIN departments d ON d.id=c.department_id WHERE c.active=1 ORDER BY d.name,c.code,c.title")
        if all_c:
            chosen=st.selectbox("Course version",all_c,format_func=lambda x:f"{x['department']} · {x['code']} — {x['title']}")
            with st.form("edit_course"):
                ec=st.text_input("Course code",chosen['code']); et=st.text_input("Course title",chosen['title']); ed=st.selectbox("Department",[d['name'] for d in deps],index=[d['name'] for d in deps].index(chosen['department'])); off=st.checkbox("Deactivate this course version")
                if st.form_submit_button("Save course"):
                    if not ec.strip() or not et.strip(): st.error("Course code and title are mandatory.")
                    else:
                        run("UPDATE courses SET code=?,title=?,department_id=?,active=? WHERE id=?",(ec.strip().upper(),et.strip(),depmap[ed],0 if off else 1,chosen['id'])); st.success("Course updated."); st.rerun()

else: # Reports: audit-entry input + teacher-wise Excel export
    st.header("Reports")
    st.write("First select the bill-owning degree department and semester/examination. Then enter each audit-sheet duty line and quantity. Download separate Excel bills for each teacher or all bills as a ZIP.")
    st.info("Minimum/maximum rules are applied to each entered duty line. For items whose official rule is per section, per course, or per year-semester, enter one line per applicable scope; scrutiny must have no more than two scrutinizers in the year-semester. Shared fees must be split into the actual share per teacher before entering the quantity/rate line. Review the breakdown before official submission.")
    deps=rows("SELECT id,name FROM departments ORDER BY name"); depmap={d['name']:d['id'] for d in deps}
    with st.form("exam_header"):
        c1,c2,c3=st.columns(3)
        with c1: deptname=st.selectbox("Department",[d['name'] for d in deps]); degree=st.selectbox("Degree",["B. Sc. Engg.","B. Sc.","B. Architecture","Other"])
        with c2: year_sem=st.text_input("Year / Semester",placeholder="1st Year 1st Semester"); exam_name=st.text_input("Examination name",placeholder="Examination"); exam_year=st.number_input("Examination year",min_value=2000,max_value=2100,value=datetime.now().year)
        with c3: exam_type=st.selectbox("Examination type",["Regular","Review","Special"]); audit_ref=st.text_input("Audit sheet reference (optional)")
        save_header=st.form_submit_button("Save / select semester audit",type="primary")
    if save_header:
        if not year_sem.strip() or not exam_name.strip(): st.error("Year/Semester and Examination name are required.")
        else:
            exam_key=f"{depmap[deptname]}|{degree}|{year_sem.strip()}|{exam_name.strip()}|{int(exam_year)}|{exam_type}"
            st.session_state['exam_bill_current']={"exam_key":exam_key,"department_id":depmap[deptname],"department":deptname,"degree":degree,"year_semester":year_sem.strip(),"exam_name":exam_name.strip(),"exam_year":int(exam_year),"exam_type":exam_type,"audit_ref":audit_ref.strip()}
            st.success("Semester audit selected. Add duty lines below.")
    current=st.session_state.get('exam_bill_current')
    if current:
        st.subheader(f"Audit sheet: {current['department']} · {current['year_semester']} · {current['exam_name']} ({current['exam_year']} / {current['exam_type']})")
        teacherlist=rows("SELECT t.id,t.name,t.designation,d.name home_department FROM teachers t JOIN departments d ON d.id=t.department_id WHERE t.active=1 ORDER BY d.name,t.name")
        ratelist=rows("SELECT id,item,rate,unit,minimum,maximum FROM rates WHERE active=1 ORDER BY item")
        courselist=rows("SELECT c.id,c.code,c.title,d.name course_department FROM courses c JOIN departments d ON d.id=c.department_id WHERE c.active=1 ORDER BY d.name,c.code,c.title")
        if not teacherlist: st.warning("No teachers assigned to this department yet. Go to Teachers and assign the relevant teachers.")
        if not ratelist: st.warning("No active bill rates. Check Bill Rates menu.")
        if teacherlist and ratelist:
            with st.form("add_duty_line"):
                a,b=st.columns(2)
                with a: teacher=st.selectbox("Teacher",teacherlist,format_func=lambda x:f"{x['name']} — {x['home_department'] or 'Home department unverified'}"); rate=st.selectbox("Bill item",ratelist,format_func=lambda x:f"{x['item']} — Tk {x['rate']:,.2f} / {x['unit']}")
                with b: course=st.selectbox("Course (optional)",[None]+courselist,format_func=lambda x:"Not course-specific" if x is None else f"{x['course_department']} · {x['code']} — {x['title']}"); qty=st.number_input("Eligible quantity (scripts/students/hours/shifts, per rate unit)",min_value=0.0,value=0.0,step=1.0)
                scope=st.text_input("Calculation scope (required)",value=current['year_semester'],help="Use the same scope text for lines sharing one limit. Answer-script examination: Section A/B; scrutiny/moderation: year-semester; sessional: course code; supervision: Supervisor overall.")
                desc=st.text_input("Description / audit-sheet note",value=rate['item']); rate_override=st.number_input("Rate override (BDT; leave equal to saved rate)",min_value=0.0,value=float(rate['rate']),step=1.0)
                add=st.form_submit_button("Add audit-sheet line")
            if add:
                if not scope.strip(): st.error("Calculation scope is required."); st.stop()
                existing=rows("SELECT id FROM bills WHERE exam_key=? AND teacher_id=?",(current['exam_key'],teacher['id']))
                if "scrutiny" in rate['item'].lower():
                    claimed=rows("SELECT DISTINCT b.teacher_id FROM bills b JOIN bill_lines l ON l.bill_id=b.id WHERE b.department_id=? AND b.degree=? AND b.year_semester=? AND b.exam_year=? AND l.rate_id=?",(current['department_id'],current['degree'],current['year_semester'],current['exam_year'],rate['id']))
                    if len(claimed)>=2 and not any(int(x['teacher_id'])==int(teacher['id']) for x in claimed):
                        st.error("The official rule allows a maximum of two scrutinizers for this year-semester."); st.stop()
                if existing: bid=existing[0]['id']
                else: bid=run("INSERT INTO bills(exam_key,department_id,teacher_id,degree,year_semester,exam_name,exam_year,exam_type,created_at) VALUES(?,?,?,?,?,?,?,?,?)",(current['exam_key'],current['department_id'],teacher['id'],current['degree'],current['year_semester'],current['exam_name'],current['exam_year'],current['exam_type'],datetime.now().isoformat(timespec='seconds')))
                amount=calc_amount(qty,rate_override,0,0)
                run("INSERT INTO bill_lines(bill_id,rate_id,course_id,description,quantity,rate,amount,scope_key) VALUES(?,?,?,?,?,?,?,?)",(bid,rate['id'],course['id'] if course else None,desc.strip() or rate['item'],float(qty),float(rate_override),amount,scope.strip().casefold()))
                recalculate_scope(bid,rate['id'],scope.strip().casefold(),float(rate['minimum']),float(rate['maximum']))
                refreshed=rows("SELECT amount FROM bill_lines WHERE bill_id=? AND rate_id=? AND scope_key=? ORDER BY id DESC LIMIT 1",(bid,rate['id'],scope.strip().casefold()))
                st.success(f"Line added. Amount after scope limits: Tk {refreshed[0]['amount']:,.2f}"); st.rerun()
        bills=rows("SELECT b.*,t.name teacher_name,t.designation,d.name department FROM bills b JOIN teachers t ON t.id=b.teacher_id JOIN departments d ON d.id=b.department_id WHERE b.exam_key=? ORDER BY t.name",(current['exam_key'],))
        if bills:
            st.subheader("Teacher-wise bill preview")
            summary=[]
            for bill in bills:
                lines=rows("SELECT l.*,r.item,c.code course_code,c.title course_title FROM bill_lines l JOIN rates r ON r.id=l.rate_id LEFT JOIN courses c ON c.id=l.course_id WHERE l.bill_id=? ORDER BY l.id",(bill['id'],))
                total=sum(float(x['amount']) for x in lines)
                summary.append({"Teacher":bill['teacher_name'],"Department":bill['department'],"Bill lines":len(lines),"Total (BDT)":total})
            st.dataframe(pd.DataFrame(summary),use_container_width=True,hide_index=True)
            selected=st.selectbox("Preview / download individual teacher bill",bills,format_func=lambda x:x['teacher_name'])
            lines=rows("SELECT l.*,r.item,c.code course_code,c.title course_title FROM bill_lines l JOIN rates r ON r.id=l.rate_id LEFT JOIN courses c ON c.id=l.course_id WHERE l.bill_id=? ORDER BY l.id",(selected['id'],))
            total=sum(float(x['amount']) for x in lines)
            st.metric(f"{selected['teacher_name']} — total bill",f"Tk {total:,.2f}")
            excel=make_teacher_xlsx(selected,lines)
            safe=re.sub(r"[^A-Za-z0-9_-]+","_",selected['teacher_name']).strip('_')
            st.download_button("Download this teacher's Excel bill",excel,file_name=f"{safe}_{current['exam_year']}_{current['exam_type']}_exam_bill.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",type="primary")
            if st.button("Delete all audit lines for selected teacher",type="secondary"):
                with db() as c: c.execute("DELETE FROM bills WHERE id=?",(selected['id'],)); c.commit()
                st.success("Teacher bill removed."); st.rerun()
            z=io.BytesIO()
            with zipfile.ZipFile(z,"w",zipfile.ZIP_DEFLATED) as archive:
                for bill in bills:
                    ls=rows("SELECT l.*,r.item,c.code course_code,c.title course_title FROM bill_lines l JOIN rates r ON r.id=l.rate_id LEFT JOIN courses c ON c.id=l.course_id WHERE l.bill_id=? ORDER BY l.id",(bill['id'],))
                    nm=re.sub(r"[^A-Za-z0-9_-]+","_",bill['teacher_name']).strip('_') or f"teacher_{bill['teacher_id']}"
                    archive.writestr(f"{nm}_exam_bill.xlsx",make_teacher_xlsx(bill,ls))
            st.download_button("Download ALL teacher-wise Excel bills (.zip)",z.getvalue(),file_name=f"{current['exam_year']}_{current['exam_type']}_teacher_bills.zip",mime="application/zip")
        else: st.info("No audit lines have been entered for this semester yet.")

st.divider()
st.caption("Department names and selected faculty entries are seeded from DUET's public website; remaining historical examiner names are left unassigned until verified. Course titles seeded from historical records require curriculum verification. Rates and calculation scopes must be checked against the official approved remuneration order. Streamlit Cloud local SQLite may not persist across rebuilds unless durable storage is configured.")
