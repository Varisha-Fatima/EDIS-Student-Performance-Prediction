"""
EDIS — Education Data Intelligence System
Multan University
Version 3.3 — Updated

Changes from v3.2:
  1. Overview header: removed admin/user avatar and search bar — date badge only (dynamic)
  2. Student card: removed dark tooltip/overlay block entirely — clean HTML only
  3. Logo is now clickable and redirects to home page
  4. Student ID + Name displayed on one clear line in class assessment
  5. Quick Actions labels ("Single Assessment", "Class Assessment") on single line
  6. Updated grading logic:
       Low Risk     → Grade B or A  (next target: A+)
       Moderate Risk → Grade C
       High Risk    → Grade D or F
"""

# ══════════════════════════════════════════════════════════════════════════════
# IMPORTS
# ══════════════════════════════════════════════════════════════════════════════
import streamlit as st #Web app / dashboard banane ke liye use hoti hai.
import pandas as pd #Tables, CSV files aur dataset
import numpy as np #Mathematical calculations
import joblib
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                 Table, TableStyle, HRFlowable)
from reportlab.lib.enums import TA_CENTER, TA_RIGHT
import datetime
import random
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import pdfplumber
import re
import io
import openpyxl

# ══════════════════════════════════════════════════════════════════════════════
# PAGE CONFIG
# ══════════════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="EDIS — Education Data Intelligence System",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ══════════════════════════════════════════════════════════════════════════════
# GLOBAL CSS
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&display=swap');

* { font-family: 'Plus Jakarta Sans', sans-serif !important; box-sizing: border-box; }

#MainMenu, footer, header,
[data-testid="stDeployButton"],
[data-testid="collapsedControl"],
button[kind="header"],
.stApp > header { visibility: hidden; display: none !important; }

[data-testid="stAppViewContainer"] { background: #F0F4FA !important; }
[data-testid="stHeader"] { display: none !important; }

/* ── Sidebar ── */
[data-testid="stSidebar"] {
    background: #111827 !important;
    border-right: none !important;
    width: 240px !important;
}
section[data-testid="stSidebar"] > div:first-child { padding: 0 !important; }
[data-testid="stSidebar"],
[data-testid="stSidebar"] *,
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] span,
[data-testid="stSidebar"] div,
[data-testid="stSidebar"] label {
    color: #CBD5E1 !important;
    -webkit-text-fill-color: #CBD5E1 !important;
}
[data-testid="stSidebar"] .stRadio div[role="radiogroup"] label {
    background: transparent !important; border-radius: 8px !important;
    padding: 10px 16px !important; margin-bottom: 2px !important;
    border: none !important; color: #94A3B8 !important;
    -webkit-text-fill-color: #94A3B8 !important; font-weight: 500 !important;
    font-size: 13.5px !important; display: flex !important;
    align-items: center !important; gap: 10px !important;
    transition: all 0.15s !important; cursor: pointer !important;
}
[data-testid="stSidebar"] .stRadio div[role="radiogroup"] label:hover {
    background: rgba(255,255,255,0.07) !important;
    color: #F1F5F9 !important; -webkit-text-fill-color: #F1F5F9 !important;
}
[data-testid="stSidebar"] .stRadio div[role="radiogroup"] label p {
    color: inherit !important; -webkit-text-fill-color: inherit !important;
}
[data-testid="stSidebar"] [aria-checked="true"] {
    background: #2563EB !important; color: #FFFFFF !important;
    -webkit-text-fill-color: #FFFFFF !important; border-radius: 8px !important;
}
[data-testid="stSidebar"] [data-testid="stMetricLabel"] p {
    color: #64748B !important; -webkit-text-fill-color: #64748B !important; font-size: 11px !important;
}
[data-testid="stSidebar"] [data-testid="stMetricValue"] {
    color: #FFFFFF !important; -webkit-text-fill-color: #FFFFFF !important;
    font-weight: 800 !important; font-size: 22px !important;
}
[data-testid="stSidebar"] [data-testid="stMetric"] {
    background: rgba(255,255,255,0.04) !important;
    border: 1px solid rgba(255,255,255,0.08) !important;
    border-radius: 10px !important; padding: 12px 14px !important;
}
[data-testid="stSidebar"] hr { border-color: rgba(255,255,255,0.08) !important; margin: 8px 0 !important; }

/* ── Main text ── */
[data-testid="stAppViewContainer"] p,
[data-testid="stAppViewContainer"] span,
[data-testid="stAppViewContainer"] div { color: #1E293B; }
[data-testid="stAppViewContainer"] h1,
[data-testid="stAppViewContainer"] h2,
[data-testid="stAppViewContainer"] h3 {
    color: #0F172A !important; font-family: 'Plus Jakarta Sans', sans-serif !important;
}

/* ── Metric cards ── */
[data-testid="stMetric"] {
    background: white !important; border: 1px solid #E2E8F0 !important;
    border-radius: 14px !important; padding: 20px 22px !important;
    box-shadow: 0 1px 8px rgba(15,27,53,0.06) !important;
}
[data-testid="stMetricLabel"] { color: #64748B !important; font-size:12px !important; font-weight:500 !important; }
[data-testid="stMetricValue"] { color: #0F172A !important; font-weight:800 !important; font-size:26px !important; }
[data-testid="stMetricDelta"] { font-size:11px !important; }

.section-title {
    font-size: 12px; font-weight: 700; letter-spacing: 0.1em;
    text-transform: uppercase; color: #2563EB !important;
    padding-bottom: 8px; margin-bottom: 16px; border-bottom: 2px solid #DBEAFE;
}

/* ── Risk cards ── */
.risk-high {
    background: #FFF5F5; border: 1px solid #FECACA; border-left: 4px solid #EF4444;
    border-radius: 14px; padding: 20px 24px; margin: 10px 0;
    box-shadow: 0 4px 16px rgba(239,68,68,0.10);
}
.risk-medium {
    background: #FFFBEB; border: 1px solid #FDE68A; border-left: 4px solid #F59E0B;
    border-radius: 14px; padding: 20px 24px; margin: 10px 0;
    box-shadow: 0 4px 16px rgba(245,158,11,0.10);
}
.risk-low {
    background: #F0FDF4; border: 1px solid #BBF7D0; border-left: 4px solid #22C55E;
    border-radius: 14px; padding: 20px 24px; margin: 10px 0;
    box-shadow: 0 4px 16px rgba(34,197,94,0.10);
}
.attendance-warning {
    background: #FFF1F2; border: 1px solid #FECDD3; border-left: 4px solid #F43F5E;
    border-radius: 14px; padding: 16px 20px; margin: 10px 0;
    box-shadow: 0 4px 16px rgba(244,63,94,0.12);
}
.email-section {
    background: #EFF6FF; border: 1px solid #BFDBFE; border-left: 4px solid #3B82F6;
    border-radius: 14px; padding: 18px 22px; margin: 10px 0;
}
.pdf-import-section {
    background: #F0FDF4; border: 1px solid #BBF7D0; border-left: 4px solid #22C55E;
    border-radius: 14px; padding: 18px 22px; margin: 10px 0;
}
.student-card {
    background: white; border: 1px solid #E2E8F0; border-radius: 12px;
    padding: 16px 20px; margin: 10px 0; box-shadow: 0 1px 6px rgba(15,27,53,0.05);
}
.lab-sub-section {
    background: #EFF6FF; border: 1px solid #BFDBFE;
    border-radius: 10px; padding: 14px 18px; margin: 8px 0;
}
.suggestion-box {
    background: #FFFBEB; border: 1px solid #FDE68A; border-left: 4px solid #F59E0B;
    border-radius: 14px; padding: 16px 20px; margin: 10px 0;
}

/* ── Inputs ── */
input, textarea,
[data-testid="stTextInput"] input,
[data-testid="stNumberInput"] input,
[data-testid="stTextArea"] textarea {
    background: #F8FAFC !important; color: #1E293B !important;
    border: 1.5px solid #E2E8F0 !important; border-radius: 8px !important;
    font-size: 13.5px !important;
}
input:focus, textarea:focus {
    border-color: #3B82F6 !important;
    box-shadow: 0 0 0 3px rgba(59,130,246,0.12) !important; outline: none !important;
}
[data-testid="stSelectbox"] div[data-baseweb="select"] > div {
    background: #F8FAFC !important; border: 1.5px solid #E2E8F0 !important;
    border-radius: 8px !important; color: #1E293B !important;
}
[data-testid="stSelectbox"] * { color: #1E293B !important; }
[data-testid="stTextInput"] label,
[data-testid="stNumberInput"] label,
[data-testid="stTextArea"] label,
[data-testid="stSelectbox"] label,
[data-testid="stSlider"] label,
[data-testid="stCheckbox"] label {
    color: #374151 !important; font-weight: 600 !important; font-size: 12.5px !important;
}
[data-testid="stCheckbox"] span { color: #374151 !important; }
[data-testid="stSlider"] * { color: #1E293B !important; }
[data-testid="stCaptionContainer"] p { color: #94A3B8 !important; font-size: 11.5px !important; }

/* ── Buttons ── */
.stButton > button {
    background: linear-gradient(135deg,#1D4ED8,#3B82F6) !important;
    color: white !important; border: none !important;
    border-radius: 9px !important; padding: 12px 24px !important;
    font-size: 14px !important; font-weight: 600 !important; width: 100% !important;
    box-shadow: 0 3px 12px rgba(29,78,216,0.30) !important; transition: all 0.18s !important;
}
.stButton > button:hover {
    transform: translateY(-1px) !important;
    box-shadow: 0 5px 18px rgba(29,78,216,0.40) !important;
}
[data-testid="stDownloadButton"] > button {
    background: linear-gradient(135deg,#059669,#10B981) !important;
    color: white !important; border: none !important;
    border-radius: 9px !important; font-weight: 600 !important; width: 100% !important;
}
[data-testid="stDataFrame"] {
    border-radius: 12px !important; overflow: hidden !important;
    box-shadow: 0 1px 8px rgba(15,27,53,0.06) !important; border: 1px solid #E2E8F0 !important;
}
[data-testid="stAlert"] { border-radius: 10px !important; }
hr { border-color: #E2E8F0 !important; }

.page-header {
    background: white; border-radius: 16px; padding: 24px 28px; margin-bottom: 20px;
    border: 1px solid #E2E8F0; box-shadow: 0 1px 8px rgba(15,27,53,0.05);
}
.page-header h2 {
    font-size: 22px !important; font-weight: 800 !important;
    color: #0F172A !important; margin-bottom: 4px !important;
}
.page-header p { color: #64748B !important; font-size: 13.5px !important; margin: 0 !important; }

.stat-card {
    background: white; border: 1px solid #E2E8F0; border-radius: 14px;
    padding: 20px 22px; box-shadow: 0 1px 8px rgba(15,27,53,0.05);
    position: relative; overflow: hidden;
}
.stat-icon {
    width: 44px; height: 44px; border-radius: 10px;
    display: flex; align-items: center; justify-content: center;
    font-size: 20px; margin-bottom: 12px;
}
.stat-value { font-size: 28px; font-weight: 800; color: #0F172A; line-height: 1; margin-bottom: 4px; }
.stat-label { font-size: 12px; font-weight: 600; color: #64748B; margin-bottom: 2px; }
.stat-sub   { font-size: 11px; color: #94A3B8; }
.stat-bar   { height: 3px; border-radius: 2px; margin-top: 14px; background: #F1F5F9; }
.stat-bar-fill { height: 100%; border-radius: 2px; }

/* ── Quick Action Cards ── */
.quick-action-card {
    background: white; border: 1px solid #E2E8F0; border-radius: 12px;
    padding: 14px 18px; display: flex; align-items: center; gap: 12px;
    margin-bottom: 8px; box-shadow: 0 1px 4px rgba(15,27,53,0.04);
    transition: box-shadow 0.15s, transform 0.15s; cursor: pointer;
}
.quick-action-card:hover {
    box-shadow: 0 4px 12px rgba(15,27,53,0.10); transform: translateY(-1px);
}
.quick-action-icon {
    border-radius: 10px; width: 38px; height: 38px;
    display: flex; align-items: center; justify-content: center;
    font-size: 17px; flex-shrink: 0;
}
.quick-action-text { display: flex; flex-direction: column; gap: 1px; }
.quick-action-title { font-weight: 700; font-size: 13px; color: #0F172A; white-space: nowrap; }
.quick-action-sub   { font-size: 11.5px; color: #94A3B8; white-space: nowrap; }

/* ── Logo link ── */
.logo-link { cursor: pointer; text-decoration: none; display: inline-flex; }
.logo-link:hover { opacity: 0.85; }

/* ── Tabs ── */
[data-testid="stTabs"] [role="tablist"] {
    background: #F1F5F9 !important; border-radius: 10px !important;
    padding: 4px !important; gap: 4px !important; border-bottom: none !important;
}
[data-testid="stTabs"] [role="tab"] {
    border-radius: 8px !important; padding: 8px 18px !important;
    font-weight: 600 !important; font-size: 13px !important;
    color: #64748B !important; border: none !important; background: transparent !important;
}
[data-testid="stTabs"] [role="tab"][aria-selected="true"] {
    background: white !important; color: #1D4ED8 !important;
    -webkit-text-fill-color: #1D4ED8 !important;
    box-shadow: 0 1px 4px rgba(15,27,53,0.10) !important;
}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# CONSTANTS
# ══════════════════════════════════════════════════════════════════════════════
MOTIVATIONAL_QUOTES = [
    ("The secret of getting ahead is getting started.", "Mark Twain"),
    ("It always seems impossible until it's done.", "Nelson Mandela"),
    ("Believe you can and you're halfway there.", "Theodore Roosevelt"),
    ("Success is the sum of small efforts repeated day in and day out.", "Robert Collier"),
    ("Education is the most powerful weapon you can use to change the world.", "Nelson Mandela"),
    ("The expert in anything was once a beginner.", "Helen Hayes"),
    ("Push yourself because no one else is going to do it for you.", "Anonymous"),
    ("Great things never come from comfort zones.", "Anonymous"),
    ("Dream it. Wish it. Do it.", "Anonymous"),
    ("The harder you work for something, the greater you'll feel when you achieve it.", "Anonymous"),
]

ISSUE_SUGGESTIONS = {
    "Select an issue...": "",
    "Attendance Issue": (
        "The student's attendance is critically below the required threshold of 75%. As per university policy, "
        "a student with attendance below 75% is NOT ELIGIBLE to sit in the final examination. Consistent class "
        "attendance is essential for understanding course material and maintaining participation marks. It is "
        "strongly recommended that the student prioritise attending all remaining sessions and seek permission "
        "for any unavoidable absences."
    ),
    "Quiz Performance Weak": (
        "The student's quiz scores indicate significant gaps in conceptual understanding. Regular revision of "
        "lecture notes, attempting practice quizzes, and actively seeking clarification from the course teacher "
        "during office hours will help improve performance considerably."
    ),
    "Assignment Weak": (
        "Assignment submissions have been inconsistent or below the required academic standard. Timely completion "
        "and thorough effort on assignments directly impacts continuous assessment marks. The student is advised "
        "to manage time effectively and seek guidance when needed."
    ),
    "Mid-Term Weak": (
        "The mid-term examination result suggests difficulty with core subject material. A structured revision "
        "plan using the course textbook, past papers, and collaborative study groups is strongly recommended "
        "before the final examination."
    ),
    "Participation Low": (
        "Class participation has been minimal throughout the semester. Active engagement in discussions, asking "
        "questions, and contributing to group activities enhances both conceptual understanding and assessment "
        "marks. The student is encouraged to take a more active role."
    ),
    "Lab Performance Weak": (
        "Laboratory work performance is below the expected standard. Revisiting lab manuals, practicing "
        "experimental procedures, and seeking timely assistance from the lab instructor will improve both "
        "understanding and scores significantly."
    ),
    "Project Incomplete": (
        "The project submission is incomplete or does not meet the required academic standard. The student must "
        "prioritise completing all project components, follow the rubric carefully, and consult the teacher for "
        "guidance before the submission deadline."
    ),
    "Overall At-Risk": (
        "The student is currently at high risk of underperforming in the final examination across multiple "
        "assessment components. An urgent meeting with the course teacher and academic advisor is strongly "
        "recommended to develop a personalised and targeted improvement plan."
    ),
    "Good Performance": (
        "The student is performing commendably across all assessment components. Maintaining current study "
        "discipline, continuing active class participation, and revising consistently will help achieve "
        "excellent results in the final examination."
    ),
}

# ══════════════════════════════════════════════════════════════════════════════
# RISK CLASSIFICATION
# ══════════════════════════════════════════════════════════════════════════════
def compute_overall_pct(marks_data: list) -> float:
    scoreable = [m for m in marks_data if m['label'] != 'Attendance']
    obtained  = sum(m['obtained'] for m in scoreable)
    possible  = sum(m['max']      for m in scoreable)
    return (obtained / possible * 100) if possible > 0 else 0.0


def check_attendance_eligibility(marks_data: list):
    for m in marks_data:
        if m['label'] == 'Attendance':
            pct = (m['obtained'] / m['max'] * 100) if m['max'] > 0 else 0
            return pct < 75, pct
    return False, 100.0


def classify_risk(overall_pct: float, att_pct: float = 100.0):
    if att_pct < 75:
        return 'Low', 'High Risk', f'{overall_pct:.1f}%'
    if overall_pct >= 60:
        return 'High', 'Low Risk', f'{overall_pct:.1f}%'
    elif overall_pct >= 40:
        return 'Medium', 'Moderate Risk', f'{overall_pct:.1f}%'
    else:
        return 'Low', 'High Risk', f'{overall_pct:.1f}%'


def predict_and_classify(marks_data, *_args):
    att_ineligible, att_pct = check_attendance_eligibility(marks_data)
    overall_pct = compute_overall_pct(marks_data)
    label, _display, risk_pct_str = classify_risk(overall_pct, att_pct)
    return label, risk_pct_str, overall_pct, att_ineligible, att_pct


def risk_display(label: str) -> str:
    return {"Low": "High Risk", "Medium": "Moderate Risk", "High": "Low Risk"}.get(label, label)


# ══════════════════════════════════════════════════════════════════════════════
# UPDATED GRADE ANALYSER
# Low Risk  → Grade B or A   (next: A+)
# Moderate  → Grade C
# High Risk → Grade D or F
# ══════════════════════════════════════════════════════════════════════════════
def analyze_grade(marks_data: list):
    scoreable = [m for m in marks_data if m['label'] != 'Attendance']
    obtained  = sum(m['obtained'] for m in scoreable)
    possible  = sum(m['max']      for m in scoreable)
    pct = (obtained / possible * 100) if possible > 0 else 0

    att_ineligible, att_pct = check_attendance_eligibility(marks_data)
    overall_pct = compute_overall_pct(marks_data)
    risk_label, _, _ = classify_risk(overall_pct, att_pct)

    # Grading logic tied to risk level
    if risk_label == 'High':          # Low Risk → B or A
        if pct >= 85:
            grade, next_g, needed = 'A', 'A+', 95
        else:
            grade, next_g, needed = 'B', 'A', 85
    elif risk_label == 'Medium':      # Moderate Risk → C
        grade, next_g, needed = 'C', 'B', 60
    else:                             # High Risk → D or F
        if pct >= 40:
            grade, next_g, needed = 'D', 'C', 50
        else:
            grade, next_g, needed = 'F', 'D', 40

    suggestions = []
    if needed:
        for m in scoreable:
            cp = (m['obtained'] / m['max'] * 100) if m['max'] > 0 else 0
            if cp < 70:
                suggestions.append({
                    'area': m['label'], 'current': round(cp, 1),
                    'potential': round(m['max'] - m['obtained'], 1),
                    'priority': 'High' if cp < 50 else 'Medium'
                })
    suggestions.sort(key=lambda x: x['current'])
    return grade, next_g, round(pct, 1), suggestions


# ══════════════════════════════════════════════════════════════════════════════
# AUTO-ISSUE DETECTION
# ══════════════════════════════════════════════════════════════════════════════
def auto_detect_issue(marks_data: list, label: str) -> str:
    att_ineligible, _ = check_attendance_eligibility(marks_data)
    if att_ineligible:
        return "Attendance Issue"
    if label == 'High':
        return "Good Performance"
    issues = []
    for m in marks_data:
        lbl = m['label']
        pct = (m['obtained'] / m['max'] * 100) if m['max'] > 0 else 0
        if lbl == 'Attendance' and pct < 75:
            issues.append(('Attendance Issue', pct))
        elif 'Quiz' in lbl and pct < 60:
            issues.append(('Quiz Performance Weak', pct))
        elif 'Assignment' in lbl and pct < 60:
            issues.append(('Assignment Weak', pct))
        elif lbl == 'Mid-Term' and pct < 60:
            issues.append(('Mid-Term Weak', pct))
        elif lbl == 'Participation' and pct < 60:
            issues.append(('Participation Low', pct))
        elif 'Lab' in lbl and pct < 60:
            issues.append(('Lab Performance Weak', pct))
        elif lbl == 'Project' and pct < 60:
            issues.append(('Project Incomplete', pct))
    if label == 'Low' and len(issues) >= 2:
        return "Overall At-Risk"
    if issues:
        issues.sort(key=lambda x: x[1])
        return issues[0][0]
    if label == 'Low':
        return "Overall At-Risk"
    if label == 'Medium':
        return "Mid-Term Weak"
    return "Select an issue..."


# ══════════════════════════════════════════════════════════════════════════════
# EMAIL
# ══════════════════════════════════════════════════════════════════════════════
SENDER_EMAIL    = "v09082350@gmail.com"
SENDER_PASSWORD = "hdtvbyvilbeaepen"


def send_email_with_pdf(to_email, student_name, risk_label, subject_name, pdf_buffer, suggestion_text):
    try:
        risk_map = {
            "Low":    ("HIGH RISK",     "Immediate academic intervention is recommended."),
            "Medium": ("MODERATE RISK", "Targeted support and close monitoring advised."),
            "High":   ("LOW RISK",      "Keep up the great work and maintain your progress!"),
        }
        risk_title, risk_note = risk_map.get(risk_label, ("UNKNOWN", ""))
        body = f"""Dear {student_name or 'Student'},

This is an official academic performance notification from EDIS — Education Data Intelligence System, Multan University.

PERFORMANCE SUMMARY
Subject      : {subject_name or 'N/A'}
Risk Status  : {risk_title}
{risk_note}

TEACHER RECOMMENDATION
{suggestion_text.strip() if suggestion_text.strip() else 'No specific recommendation provided.'}

Please find your detailed academic report attached to this email.
If you have any questions, please contact your course teacher or academic advisor.

Best regards,
EDIS System — Education Data Intelligence System — Multan University
(This is an automated message — please do not reply directly.)
"""
        msg = MIMEMultipart()
        msg['From']    = SENDER_EMAIL
        msg['To']      = to_email
        msg['Subject'] = f"EDIS — Academic Report | {student_name or 'Student'} | {risk_title}"
        msg.attach(MIMEText(body, 'plain'))
        pdf_buffer.seek(0)
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(pdf_buffer.read())
        encoders.encode_base64(part)
        safe_name = (student_name or "Student").replace(" ", "_")
        part.add_header('Content-Disposition', f'attachment; filename="EDIS_Report_{safe_name}.pdf"')
        msg.attach(part)
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.send_message(msg)
        server.quit()
        return True, ""
    except smtplib.SMTPAuthenticationError:
        return False, "Authentication failed. Please check Gmail credentials / App Password."
    except smtplib.SMTPException as e:
        return False, f"SMTP error: {str(e)}"
    except Exception as e:
        return False, f"Unexpected error: {str(e)}"


# ══════════════════════════════════════════════════════════════════════════════
# DATA & MODEL LOADING
# ══════════════════════════════════════════════════════════════════════════════
@st.cache_resource
def load_model():
    try:
        return joblib.load('edupredict_model.pkl'), joblib.load('label_encoder.pkl')
    except FileNotFoundError as e:
        st.error(f"Model file not found: {e}")
        st.stop()


@st.cache_data
def load_data():
    try:
        return pd.read_csv('multan_university_dataset.csv')
    except FileNotFoundError:
        st.error("Dataset not found.")
        st.stop()


model, le = load_model()
df        = load_data()


# ══════════════════════════════════════════════════════════════════════════════
# PDF STUDENT LIST PARSER
# ══════════════════════════════════════════════════════════════════════════════
_HEADER_SKIP = {
    'reg no', 'reg no name', 'name', 'student name', 'full name', 'sname',
    'roll no', 'roll number', 'registration no', 'registration number',
    'sr no', 'sno', 's.no', 'serial', 'id', 'student', 'roll', 'registration',
}

_ROLL_RE = re.compile(
    r'(\b\d{4}-[A-Z]{2,6}-[A-Z]-\d{2,5}\b'
    r'|\b\d{4}-[A-Z]{2,6}-\d{2,5}\b'
    r'|\b[A-Z]{1,4}[-/]\d{2,4}[-/][A-Z]{0,4}\d{1,4}\b'
    r'|\b[A-Z]{2,4}\d{2}[-/][A-Z]{2,4}[-/]\d{3,5}\b'
    r'|\bFA\d{2}[-/][A-Z]{2,4}[-/]\d{3,5}\b'
    r'|\b[A-Z]{2,6}\d{4,10}\b'
    r'|\b\d{10,14}\b'
    r'|\b\d{4}[-/][A-Z]{2,4}[-/]\d{2,4}\b)',
    re.IGNORECASE
)


def _is_header(name: str) -> bool:
    return name.lower().strip() in _HEADER_SKIP


def parse_students_from_pdf(uploaded_file) -> list:
    students = []
    try:
        pdf_bytes = uploaded_file.read()
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if not table or len(table) < 2:
                        continue
                    raw_headers = table[0] or []
                    headers = [
                        str(h).lower().strip().replace('\n', ' ').replace('\r', '') if h else ''
                        for h in raw_headers
                    ]
                    name_col = next(
                        (i for i, h in enumerate(headers)
                         if any(k in h for k in ['name', 'student', 'naam', 'sname', 'full'])),
                        None
                    )
                    roll_col = next(
                        (i for i, h in enumerate(headers)
                         if any(k in h for k in ['roll', 'reg', 'id', 'rno', 'r.no',
                                                   'enroll', 'serial', 'sr', 's.no',
                                                   'sno', 'number', 'no.', 'class'])),
                        None
                    )
                    if name_col is not None and roll_col is None:
                        for ci in range(len(headers)):
                            if ci != name_col:
                                roll_col = ci
                                break

                    for row in table[1:]:
                        if not row:
                            continue
                        try:
                            name = ''
                            roll = ''
                            if name_col is not None and name_col < len(row) and row[name_col]:
                                name = re.sub(r'\s+', ' ', str(row[name_col])).strip()
                            if roll_col is not None and roll_col < len(row) and row[roll_col]:
                                roll = re.sub(r'\s+', ' ', str(row[roll_col])).strip()

                            if roll.lower() in ('none', 'null', '', 'roll no', 'reg no', 'id',
                                                 'roll', 'registration', 'sr no', 'sr.no'):
                                roll = 'N/A'
                            if roll != 'N/A' and not _ROLL_RE.match(roll):
                                found = _ROLL_RE.findall(roll)
                                roll = found[0] if found else roll

                            if name and not _is_header(name):
                                students.append({'name': name, 'roll_no': roll or 'N/A'})
                        except Exception:
                            continue

                if not students:
                    text = page.extract_text() or ''
                    for raw_line in text.split('\n'):
                        line = raw_line.strip()
                        if len(line) < 4:
                            continue
                        clean = re.sub(r'^\d{1,3}[\.\)\s]+', '', line).strip()
                        rolls     = _ROLL_RE.findall(clean)
                        name_part = _ROLL_RE.sub('', clean).strip()
                        name_part = re.sub(r'[_|]+', ' ', name_part).strip()
                        name_part = re.sub(r'\s+', ' ', name_part).strip()
                        roll = rolls[0].strip() if rolls else 'N/A'
                        name = name_part if len(name_part) > 2 else ''
                        if name and not _is_header(name):
                            students.append({'name': name, 'roll_no': roll})

    except Exception as e:
        st.error(f"PDF parsing error: {e}")

    seen   = set()
    unique = []
    for s in students:
        key = (s['name'].lower().strip(), s['roll_no'].lower().strip())
        if key not in seen and s['name'].strip():
            seen.add(key)
            unique.append(s)
    return unique


# ══════════════════════════════════════════════════════════════════════════════
# EXCEL STUDENT LIST PARSER
# ══════════════════════════════════════════════════════════════════════════════
def parse_students_from_excel(uploaded_file) -> list:
    students = []
    try:
        file_bytes = uploaded_file.read()
        xf = pd.read_excel(io.BytesIO(file_bytes), engine='openpyxl')
        xf.columns = [str(c).lower().strip() for c in xf.columns]
        name_col = next(
            (c for c in xf.columns if any(k in c for k in ['name', 'student', 'naam', 'full'])),
            None
        )
        roll_col = next(
            (c for c in xf.columns if any(k in c for k in
             ['roll', 'reg', 'id', 'rno', 'r.no', 'number', 'enroll', 'serial', 'sr'])),
            None
        )
        if name_col:
            for _, row in xf.iterrows():
                name = str(row[name_col]).strip() if pd.notna(row.get(name_col)) else ''
                roll = str(row[roll_col]).strip() if roll_col and pd.notna(row.get(roll_col)) else 'N/A'
                if roll.lower() in ('nan', 'none', ''):
                    roll = 'N/A'
                if name and name.lower() not in ('nan', 'none', 'name', 'student name', ''):
                    students.append({'name': name, 'roll_no': roll})
        else:
            for _, row in xf.iterrows():
                vals = [str(v).strip() for v in row.values if pd.notna(v) and str(v).strip()]
                if len(vals) >= 2:
                    students.append({'name': vals[0], 'roll_no': vals[1]})
                elif len(vals) == 1:
                    students.append({'name': vals[0], 'roll_no': 'N/A'})
    except Exception as e:
        st.error(f"Excel parsing error: {e}")

    seen   = set()
    unique = []
    for s in students:
        key = (s['name'].lower(), s['roll_no'].lower())
        if key not in seen and s['name'].strip():
            seen.add(key)
            unique.append(s)
    return unique


# ══════════════════════════════════════════════════════════════════════════════
# PDF REPORT GENERATOR
# ══════════════════════════════════════════════════════════════════════════════
def make_pdf(student_info: dict, marks_data: list, label: str,
             risk_pct_str: str, suggestion_text: str, att_ineligible: bool = False) -> BytesIO:
    buffer = BytesIO()
    doc    = SimpleDocTemplate(buffer, pagesize=A4,
                               rightMargin=1.8*cm, leftMargin=1.8*cm,
                               topMargin=1.5*cm,   bottomMargin=1.5*cm)
    story  = []
    BLUE   = colors.HexColor("#1D4ED8")
    RED    = colors.HexColor("#EF4444")
    GREEN  = colors.HexColor("#22C55E")
    GOLD   = colors.HexColor("#F59E0B")
    ROSE   = colors.HexColor("#F43F5E")
    WHITE  = colors.white
    LIGHT  = colors.HexColor("#F8FAFF")

    rc = RED if label == 'Low' else (GOLD if label == 'Medium' else GREEN)
    rt = ("HIGH RISK — Immediate Intervention Required" if label == 'Low' else
          "MODERATE RISK — Monitoring Advised"          if label == 'Medium' else
          "PERFORMING WELL — On Track")

    hdr = Table([[
        Paragraph("<font size='18' color='white'><b>EDIS</b></font><br/>"
                  "<font size='8' color='#BFDBFE'>Education Data Intelligence System · Multan University</font>",
                  ParagraphStyle('H', fontName='Helvetica', leading=22)),
        Paragraph(f"<font size='8' color='#BFDBFE'>"
                  f"Date: {datetime.datetime.now().strftime('%d %B %Y')}<br/>"
                  f"Confidential — Academic Use Only</font>",
                  ParagraphStyle('HR', alignment=TA_RIGHT, fontName='Helvetica', leading=14))
    ]], colWidths=[11*cm, 6.4*cm])
    hdr.setStyle(TableStyle([
        ('BACKGROUND',    (0,0),(-1,-1), BLUE),
        ('TOPPADDING',    (0,0),(-1,-1), 16),
        ('BOTTOMPADDING', (0,0),(-1,-1), 16),
        ('LEFTPADDING',   (0,0),(-1,-1), 18),
        ('RIGHTPADDING',  (0,0),(-1,-1), 18),
    ]))
    story.append(hdr)
    story.append(Spacer(1, 0.3*cm))

    if att_ineligible:
        ab = Table([[Paragraph(
            "<font size='11' color='white'><b>ATTENDANCE BELOW 75% — NOT ELIGIBLE FOR FINAL EXAM</b></font>",
            ParagraphStyle('AB', alignment=TA_CENTER, fontName='Helvetica-Bold', leading=16)
        )]], colWidths=[17.4*cm])
        ab.setStyle(TableStyle([
            ('BACKGROUND',    (0,0),(-1,-1), ROSE),
            ('TOPPADDING',    (0,0),(-1,-1), 10),
            ('BOTTOMPADDING', (0,0),(-1,-1), 10),
        ]))
        story.append(ab)
        story.append(Spacer(1, 0.15*cm))

    rb = Table([[Paragraph(
        f"<font size='13' color='white'><b>{rt}  ({risk_pct_str})</b></font>",
        ParagraphStyle('RB', alignment=TA_CENTER, fontName='Helvetica', leading=18)
    )]], colWidths=[17.4*cm])
    rb.setStyle(TableStyle([
        ('BACKGROUND',    (0,0),(-1,-1), rc),
        ('TOPPADDING',    (0,0),(-1,-1), 12),
        ('BOTTOMPADDING', (0,0),(-1,-1), 12),
    ]))
    story.append(rb)
    story.append(Spacer(1, 0.3*cm))

    irows = [[
        Paragraph(f"<font size='9' color='#1D4ED8'><b>{k}</b></font>",
                  ParagraphStyle('IL', fontName='Helvetica')),
        Paragraph(f"<font size='9' color='#1E2A3A'>{v}</font>",
                  ParagraphStyle('IV', fontName='Helvetica'))
    ] for k, v in student_info.items()]
    itbl = Table(irows, colWidths=[3.5*cm, 4.8*cm])
    itbl.setStyle(TableStyle([
        ('ROWBACKGROUNDS', (0,0),(-1,-1), [WHITE, LIGHT]),
        ('GRID',          (0,0),(-1,-1), 0.4, colors.HexColor("#E3E8F0")),
        ('TOPPADDING',    (0,0),(-1,-1), 6),
        ('BOTTOMPADDING', (0,0),(-1,-1), 6),
        ('LEFTPADDING',   (0,0),(-1,-1), 10),
    ]))

    mhdr_row = [[
        Paragraph("<font size='9' color='white'><b>Assessment</b></font>",
                  ParagraphStyle('MH', fontName='Helvetica-Bold')),
        Paragraph("<font size='9' color='white'><b>Marks</b></font>",
                  ParagraphStyle('M2', fontName='Helvetica-Bold', alignment=TA_CENTER)),
        Paragraph("<font size='9' color='white'><b>Out of</b></font>",
                  ParagraphStyle('M3', fontName='Helvetica-Bold', alignment=TA_CENTER)),
        Paragraph("<font size='9' color='white'><b>Score %</b></font>",
                  ParagraphStyle('M4', fontName='Helvetica-Bold', alignment=TA_CENTER)),
    ]]
    mrows = []
    for item in marks_data:
        pct = round((item['obtained'] / item['max']) * 100, 1) if item['max'] > 0 else 0
        if item['label'] == 'Attendance' and pct < 75:
            c = '#F43F5E'
        else:
            c = '#EF4444' if pct < 50 else ('#F59E0B' if pct < 70 else '#22C55E')
        mrows.append([
            Paragraph(f"<font size='8' color='#1E2A3A'>{item['label']}</font>",
                      ParagraphStyle('ML', fontName='Helvetica')),
            Paragraph(f"<font size='8' color='#1E2A3A'>{item['obtained']}</font>",
                      ParagraphStyle('MC', alignment=TA_CENTER, fontName='Helvetica')),
            Paragraph(f"<font size='8' color='#9CA3AF'>{item['max']}</font>",
                      ParagraphStyle('MC2', alignment=TA_CENTER, fontName='Helvetica')),
            Paragraph(f"<font size='8' color='{c}'><b>{pct}%</b></font>",
                      ParagraphStyle('MC3', alignment=TA_CENTER, fontName='Helvetica-Bold')),
        ])
    mtbl = Table(mhdr_row + mrows, colWidths=[3.8*cm, 2*cm, 2*cm, 1.6*cm])
    mtbl.setStyle(TableStyle([
        ('BACKGROUND',    (0,0),(-1,0), BLUE),
        ('ROWBACKGROUNDS',(0,1),(-1,-1), [WHITE, LIGHT]),
        ('GRID',          (0,0),(-1,-1), 0.4, colors.HexColor("#E3E8F0")),
        ('TOPPADDING',    (0,0),(-1,-1), 6),
        ('BOTTOMPADDING', (0,0),(-1,-1), 6),
        ('LEFTPADDING',   (0,0),(-1,-1), 10),
    ]))

    layout = Table([[
        Table([[Paragraph("<font size='10' color='#1D4ED8'><b>Student Information</b></font>",
                          ParagraphStyle('SH', fontName='Helvetica-Bold'))], [itbl]],
              colWidths=[8.5*cm]),
        Table([[Paragraph("<font size='10' color='#1D4ED8'><b>Academic Performance (All Components)</b></font>",
                          ParagraphStyle('AH', fontName='Helvetica-Bold'))], [mtbl]],
              colWidths=[9.4*cm]),
    ]], colWidths=[8.7*cm, 9.7*cm])
    layout.setStyle(TableStyle([
        ('VALIGN',        (0,0),(-1,-1), 'TOP'),
        ('LEFTPADDING',   (0,0),(-1,-1), 0),
        ('RIGHTPADDING',  (0,0),(-1,-1), 6),
        ('TOPPADDING',    (0,0),(-1,-1), 0),
        ('BOTTOMPADDING', (0,0),(-1,-1), 0),
    ]))
    story.append(layout)
    story.append(Spacer(1, 0.3*cm))

    if suggestion_text.strip():
        story.append(Paragraph("<font size='10' color='#1D4ED8'><b>Teacher Recommendations</b></font>",
                                ParagraphStyle('TH', fontName='Helvetica-Bold')))
        story.append(Spacer(1, 0.12*cm))
        st_tbl = Table([[Paragraph(
            f"<font size='9' color='#1E2A3A'>{suggestion_text}</font>",
            ParagraphStyle('SG', fontName='Helvetica', leading=14)
        )]], colWidths=[17.4*cm])
        st_tbl.setStyle(TableStyle([
            ('BACKGROUND',    (0,0),(-1,-1), LIGHT),
            ('TOPPADDING',    (0,0),(-1,-1), 10),
            ('BOTTOMPADDING', (0,0),(-1,-1), 10),
            ('LEFTPADDING',   (0,0),(-1,-1), 14),
            ('GRID',          (0,0),(-1,-1), 0.4, colors.HexColor("#E3E8F0")),
        ]))
        story.append(st_tbl)
        story.append(Spacer(1, 0.25*cm))

    quote, author = random.choice(MOTIVATIONAL_QUOTES)
    qt = Table([[Paragraph(
        f"<font size='10' color='#1D4ED8'><i>\"{quote}\"</i></font><br/>"
        f"<font size='9' color='#6B7280'>— {author}</font>",
        ParagraphStyle('Q', fontName='Helvetica', leading=16, alignment=TA_CENTER)
    )]], colWidths=[17.4*cm])
    qt.setStyle(TableStyle([
        ('BACKGROUND',    (0,0),(-1,-1), colors.HexColor("#EFF6FF")),
        ('TOPPADDING',    (0,0),(-1,-1), 12),
        ('BOTTOMPADDING', (0,0),(-1,-1), 12),
        ('LEFTPADDING',   (0,0),(-1,-1), 18),
    ]))
    story.append(qt)
    story.append(Spacer(1, 0.2*cm))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#E3E8F0")))
    story.append(Spacer(1, 0.12*cm))
    story.append(Paragraph(
        "<font size='7.5' color='#9CA3AF'>EDIS · Education Data Intelligence System · "
        "Multan University · Confidential — For official academic use only.</font>",
        ParagraphStyle('FT', alignment=TA_CENTER, fontName='Helvetica')
    ))
    doc.build(story)
    buffer.seek(0)
    return buffer


# ══════════════════════════════════════════════════════════════════════════════
# CLASS ASSESSMENT REPORT — PDF
# ══════════════════════════════════════════════════════════════════════════════
def make_class_pdf(batch_df: pd.DataFrame, class_info: dict) -> BytesIO:
    buffer = BytesIO()
    doc    = SimpleDocTemplate(buffer, pagesize=A4,
                               rightMargin=1.2*cm, leftMargin=1.2*cm,
                               topMargin=1.2*cm,   bottomMargin=1.2*cm)
    story = []
    BLUE  = colors.HexColor("#1D4ED8")
    WHITE = colors.white
    LIGHT = colors.HexColor("#F8FAFF")
    RED   = colors.HexColor("#FEE2E2")
    AMBER = colors.HexColor("#FEF9C3")
    GREEN = colors.HexColor("#DCFCE7")

    hdr = Table([[
        Paragraph("<font size='16' color='white'><b>EDIS — Class Assessment Report</b></font><br/>"
                  "<font size='8' color='#BFDBFE'>Education Data Intelligence System · Multan University</font>",
                  ParagraphStyle('H', fontName='Helvetica', leading=22)),
        Paragraph(f"<font size='8' color='#BFDBFE'>"
                  f"Date: {datetime.datetime.now().strftime('%d %B %Y')}<br/>"
                  f"Dept: {class_info.get('department','N/A')} · Sem: {class_info.get('semester','N/A')}</font>",
                  ParagraphStyle('HR', alignment=TA_RIGHT, fontName='Helvetica', leading=14))
    ]], colWidths=[12*cm, 5.8*cm])
    hdr.setStyle(TableStyle([
        ('BACKGROUND',    (0,0),(-1,-1), BLUE),
        ('TOPPADDING',    (0,0),(-1,-1), 14),
        ('BOTTOMPADDING', (0,0),(-1,-1), 14),
        ('LEFTPADDING',   (0,0),(-1,-1), 16),
        ('RIGHTPADDING',  (0,0),(-1,-1), 16),
    ]))
    story.append(hdr)
    story.append(Spacer(1, 0.3*cm))

    display_cols = [c for c in batch_df.columns if c not in ('Email',)]
    col_headers  = [
        Paragraph(f"<font size='7' color='white'><b>{c}</b></font>",
                  ParagraphStyle(f'CH{i}', fontName='Helvetica-Bold', alignment=TA_CENTER))
        for i, c in enumerate(display_cols)
    ]
    table_data = [col_headers]
    n_cols     = len(display_cols)
    col_w      = 18.0 / max(n_cols, 1)
    col_widths = [col_w * cm] * n_cols

    for _, row in batch_df.iterrows():
        table_row = []
        for c in display_cols:
            val = row.get(c, '')
            table_row.append(Paragraph(
                f"<font size='7' color='#1E293B'>{val}</font>",
                ParagraphStyle(f'cv_{c}', fontName='Helvetica', alignment=TA_CENTER)
            ))
        table_data.append(table_row)

    tbl = Table(table_data, colWidths=col_widths, repeatRows=1)
    row_bg_cmds = [('BACKGROUND', (0,0), (-1,0), BLUE)]
    for ri, (_, row) in enumerate(batch_df.iterrows(), start=1):
        risk_val = str(row.get('Risk Level', ''))
        bg_color = RED if 'High' in risk_val else (AMBER if 'Moderate' in risk_val else GREEN)
        row_bg_cmds.append(('BACKGROUND', (0, ri), (-1, ri), bg_color))
    tbl.setStyle(TableStyle(row_bg_cmds + [
        ('GRID',          (0,0),(-1,-1), 0.4, colors.HexColor("#E3E8F0")),
        ('TOPPADDING',    (0,0),(-1,-1), 5),
        ('BOTTOMPADDING', (0,0),(-1,-1), 5),
        ('LEFTPADDING',   (0,0),(-1,-1), 4),
        ('RIGHTPADDING',  (0,0),(-1,-1), 4),
    ]))
    story.append(tbl)
    story.append(Spacer(1, 0.4*cm))

    total = len(batch_df)
    high  = len(batch_df[batch_df['Risk Level'] == 'High Risk'])
    med   = len(batch_df[batch_df['Risk Level'] == 'Moderate Risk'])
    low_r = len(batch_df[batch_df['Risk Level'] == 'Low Risk'])
    story.append(Paragraph(
        f"<font size='9' color='#1D4ED8'><b>Summary:</b></font> "
        f"<font size='9' color='#1E293B'>Total: {total} | High Risk: {high} | Moderate: {med} | Low Risk: {low_r}</font>",
        ParagraphStyle('SUM', fontName='Helvetica', leading=14)
    ))
    story.append(Spacer(1, 0.2*cm))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#E3E8F0")))
    story.append(Spacer(1, 0.1*cm))
    story.append(Paragraph(
        "<font size='7.5' color='#9CA3AF'>EDIS · Education Data Intelligence System · "
        "Multan University · Confidential — For official academic use only.</font>",
        ParagraphStyle('FT', alignment=TA_CENTER, fontName='Helvetica')
    ))
    doc.build(story)
    buffer.seek(0)
    return buffer


# ══════════════════════════════════════════════════════════════════════════════
# EXCEL REPORT GENERATOR
# ══════════════════════════════════════════════════════════════════════════════
def make_excel(data_df: pd.DataFrame) -> BytesIO:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
        data_df.to_excel(writer, sheet_name='Results', index=False)
        wb  = writer.book
        ws  = writer.sheets['Results']
        hf  = wb.add_format({'bold': True, 'bg_color': '#1D4ED8', 'font_color': 'white',
                              'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 11})
        hg  = wb.add_format({'bg_color': '#DCFCE7', 'border': 1, 'align': 'center', 'font_color': '#166534'})
        hm  = wb.add_format({'bg_color': '#FEF9C3', 'border': 1, 'align': 'center', 'font_color': '#713F12'})
        hl  = wb.add_format({'bg_color': '#FEE2E2', 'border': 1, 'align': 'center', 'font_color': '#991B1B'})
        cf  = wb.add_format({'border': 1, 'font_size': 10, 'align': 'left'})
        for ci, cn in enumerate(data_df.columns):
            ws.write(0, ci, cn, hf)
            ws.set_column(ci, ci, max(15, len(str(cn)) + 4))
        for ri in range(len(data_df)):
            risk = str(data_df.iloc[ri].get('Risk Level', ''))
            rf   = hl if 'High' in risk else (hm if 'Moderate' in risk else hg)
            for ci in range(len(data_df.columns)):
                val = data_df.iloc[ri, ci]
                ws.write(ri + 1, ci, val, rf if data_df.columns[ci] == 'Risk Level' else cf)
    buf.seek(0)
    return buf


# ══════════════════════════════════════════════════════════════════════════════
# RISK RESULT CARD
# ══════════════════════════════════════════════════════════════════════════════
def result_card(label, risk_pct_str, name, dept, subj,
                overall_pct=None, att_ineligible=False, att_pct=100):
    pct_str = f" · {overall_pct:.1f}% overall" if overall_pct is not None else ""

    if att_ineligible:
        st.markdown(f"""<div class="attendance-warning">
            <div style='display:flex;align-items:center;gap:10px;'>
                <div style='background:#FFE4E6;border-radius:10px;width:42px;height:42px;
                            display:flex;align-items:center;justify-content:center;font-size:22px;flex-shrink:0;'>🚫</div>
                <div>
                    <div style='font-size:16px;font-weight:800;color:#BE123C;'>EXAM INELIGIBLE — Attendance {att_pct:.1f}%</div>
                    <div style='font-size:12.5px;color:#E11D48;margin-top:3px;'>
                        Attendance is below the required <b>75%</b> threshold. This student is <b>NOT eligible</b>
                        to appear in the final examination as per university policy.</div>
                </div>
                <span style='margin-left:auto;background:#F43F5E;color:white;font-size:11px;font-weight:700;
                             padding:4px 12px;border-radius:20px;letter-spacing:0.05em;white-space:nowrap;'>EXAM BARRED</span>
            </div>
        </div>""", unsafe_allow_html=True)

    if label == 'Low':
        st.markdown(f"""<div class="risk-high">
            <div style='display:flex;align-items:center;gap:10px;margin-bottom:10px;'>
                <div style='background:#FEE2E2;border-radius:10px;width:42px;height:42px;
                            display:flex;align-items:center;justify-content:center;font-size:22px;'>🚨</div>
                <div>
                    <div style='font-size:17px;font-weight:800;color:#991B1B;'>HIGH RISK — {risk_pct_str}{pct_str}</div>
                    <div style='font-size:12px;color:#DC2626;margin-top:1px;'>Immediate intervention required</div>
                </div>
                <span style='margin-left:auto;background:#EF4444;color:white;font-size:11px;font-weight:700;
                             padding:4px 12px;border-radius:20px;letter-spacing:0.05em;'>ACTION NEEDED</span>
            </div>
            <div style='display:flex;gap:8px;flex-wrap:wrap;'>
                <span style='background:#FEE2E2;color:#991B1B;font-size:12.5px;font-weight:600;padding:4px 12px;border-radius:6px;'>{name}</span>
                <span style='background:#FEE2E2;color:#991B1B;font-size:12.5px;font-weight:600;padding:4px 12px;border-radius:6px;'>{dept}</span>
                <span style='background:#FEE2E2;color:#991B1B;font-size:12.5px;font-weight:600;padding:4px 12px;border-radius:6px;'>{subj}</span>
            </div>
        </div>""", unsafe_allow_html=True)
    elif label == 'Medium':
        st.markdown(f"""<div class="risk-medium">
            <div style='display:flex;align-items:center;gap:10px;margin-bottom:10px;'>
                <div style='background:#FEF9C3;border-radius:10px;width:42px;height:42px;
                            display:flex;align-items:center;justify-content:center;font-size:22px;'>⚡</div>
                <div>
                    <div style='font-size:17px;font-weight:800;color:#713F12;'>MODERATE RISK — {risk_pct_str}{pct_str}</div>
                    <div style='font-size:12px;color:#D97706;margin-top:1px;'>Monitoring and support required</div>
                </div>
                <span style='margin-left:auto;background:#F59E0B;color:white;font-size:11px;font-weight:700;
                             padding:4px 12px;border-radius:20px;letter-spacing:0.05em;'>MONITORING</span>
            </div>
            <div style='display:flex;gap:8px;flex-wrap:wrap;'>
                <span style='background:#FEF9C3;color:#713F12;font-size:12.5px;font-weight:600;padding:4px 12px;border-radius:6px;'>{name}</span>
                <span style='background:#FEF9C3;color:#713F12;font-size:12.5px;font-weight:600;padding:4px 12px;border-radius:6px;'>{dept}</span>
                <span style='background:#FEF9C3;color:#713F12;font-size:12.5px;font-weight:600;padding:4px 12px;border-radius:6px;'>{subj}</span>
            </div>
        </div>""", unsafe_allow_html=True)
    else:
        st.markdown(f"""<div class="risk-low">
            <div style='display:flex;align-items:center;gap:10px;margin-bottom:10px;'>
                <div style='background:#DCFCE7;border-radius:10px;width:42px;height:42px;
                            display:flex;align-items:center;justify-content:center;font-size:22px;'>✅</div>
                <div>
                    <div style='font-size:17px;font-weight:800;color:#14532D;'>LOW RISK — {risk_pct_str}{pct_str}</div>
                    <div style='font-size:12px;color:#16A34A;margin-top:1px;'>Student is on track</div>
                </div>
                <span style='margin-left:auto;background:#22C55E;color:white;font-size:11px;font-weight:700;
                             padding:4px 12px;border-radius:20px;letter-spacing:0.05em;'>ON TRACK</span>
            </div>
            <div style='display:flex;gap:8px;flex-wrap:wrap;'>
                <span style='background:#DCFCE7;color:#14532D;font-size:12.5px;font-weight:600;padding:4px 12px;border-radius:6px;'>{name}</span>
                <span style='background:#DCFCE7;color:#14532D;font-size:12.5px;font-weight:600;padding:4px 12px;border-radius:6px;'>{dept}</span>
                <span style='background:#DCFCE7;color:#14532D;font-size:12.5px;font-weight:600;padding:4px 12px;border-radius:6px;'>{subj}</span>
            </div>
        </div>""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# EDIS LOGO SVG
# ══════════════════════════════════════════════════════════════════════════════
EDIS_LOGO_SVG = """
<svg width="44" height="44" viewBox="0 0 44 44" fill="none" xmlns="http://www.w3.org/2000/svg">
  <circle cx="22" cy="22" r="22" fill="#F97316"/>
  <polygon points="22,10 34,16 22,22 10,16" fill="white" opacity="0.95"/>
  <path d="M10 16 L10 23 Q10 25 12 26 L22 30 L32 26 Q34 25 34 23 L34 16"
        fill="none" stroke="white" stroke-width="1.5" opacity="0.85"/>
  <rect x="14" y="31" width="3" height="5" rx="1" fill="white" opacity="0.9"/>
  <rect x="19" y="28" width="3" height="8" rx="1" fill="white" opacity="0.9"/>
  <rect x="24" y="25" width="3" height="11" rx="1" fill="white" opacity="0.9"/>
  <rect x="29" y="22" width="3" height="14" rx="1" fill="white" opacity="0.9"/>
</svg>
"""

# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR — Logo is now a clickable home link via query params
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    # Clickable logo that navigates to the Overview page
    st.markdown(f"""
    <div style='padding:24px 20px 12px 20px;'>
        <a href='/' target='_self' style='text-decoration:none;display:flex;align-items:center;
                  gap:12px;margin-bottom:20px;cursor:pointer;'>
            <div style='flex-shrink:0;'>{EDIS_LOGO_SVG}</div>
            <div>
                <div style='font-size:20px;font-weight:800;color:#FFFFFF;-webkit-text-fill-color:#FFFFFF;
                            letter-spacing:0.01em;font-family:"Plus Jakarta Sans",sans-serif;line-height:1.1;'>EDIS</div>
                <div style='font-size:10px;color:#94A3B8;-webkit-text-fill-color:#94A3B8;
                            font-weight:500;margin-top:2px;line-height:1.3;'>Education Data<br>Intelligence System</div>
            </div>
        </a>
        <div style='height:1px;background:rgba(255,255,255,0.07);margin-bottom:16px;'></div>
    </div>
    """, unsafe_allow_html=True)

    page = st.radio("", [
        "🏠  Overview",
        "🔍  Single Assessment",
        "👥  Class Assessment",
        "📊  Department Insights",
        "⚠️  At-Risk Registry",
    ])

    st.markdown("""
    <div style='padding:4px 20px 8px 20px;margin-top:4px;'>
        <div style='height:1px;background:rgba(255,255,255,0.07);margin-bottom:16px;'></div>
        <div style='font-size:10px;font-weight:700;letter-spacing:0.12em;text-transform:uppercase;
                    color:#475569;-webkit-text-fill-color:#475569;margin-bottom:10px;'>QUICK STATS</div>
    </div>
    """, unsafe_allow_html=True)

    total_s = len(df)
    risk_s  = len(df[df['performance_label'] == 'Low'])
    st.metric("Total Students",   f"{total_s:,}")
    st.metric("At-Risk Students", f"{risk_s:,}",
              delta=f"{risk_s / total_s * 100:.1f}%", delta_color="inverse")

    st.markdown(f"""
    <div style='padding:16px 20px 0 20px;'>
        <div style='height:1px;background:rgba(255,255,255,0.07);margin-bottom:14px;'></div>
        <div style='background:linear-gradient(135deg,rgba(29,78,216,0.35),rgba(249,115,22,0.25));
                    border:1px solid rgba(255,255,255,0.08);border-radius:10px;padding:12px 14px;'>
            <div style='font-size:11px;color:#94A3B8;-webkit-text-fill-color:#94A3B8;font-weight:700;'>TOTAL STUDENTS</div>
            <div style='font-size:26px;font-weight:800;color:#FFFFFF;-webkit-text-fill-color:#FFFFFF;line-height:1;'>{total_s:,}</div>
            <div style='font-size:11px;color:#94A3B8;-webkit-text-fill-color:#94A3B8;margin-top:2px;'>Enrolled</div>
            <div style='margin-top:8px;padding-top:8px;border-top:1px solid rgba(255,255,255,0.07);'>
                <div style='font-size:11px;color:#FCA5A5;-webkit-text-fill-color:#FCA5A5;font-weight:700;'>At-Risk Students</div>
                <div style='font-size:18px;font-weight:800;color:#F87171;-webkit-text-fill-color:#F87171;'>
                    {risk_s:,} <span style='font-size:12px;'>↑ {risk_s / total_s * 100:.1f}%</span></div>
            </div>
        </div>
        <div style='font-size:11px;color:#475569;-webkit-text-fill-color:#475569;text-align:center;margin-top:12px;'>EDIS v3.3 · Rule-Based ML</div>
        <div style='font-size:10px;color:#334155;-webkit-text-fill-color:#334155;text-align:center;margin-top:2px;'>© 2025 Multan University</div>
    </div>
    """, unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE 1 — OVERVIEW
# Clean header: Logo + title left, current date right only (no search bar, no admin avatar)
# ══════════════════════════════════════════════════════════════════════════════
if page == "🏠  Overview":
    # Dynamic date — updates on every load
    today_str = datetime.datetime.now().strftime("%A, %B %d, %Y")

    st.markdown(f"""
    <div style='background:white;border-radius:16px;padding:24px 28px;margin-bottom:20px;
                border:1px solid #E2E8F0;box-shadow:0 1px 8px rgba(15,27,53,0.05);'>
        <div style='display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:12px;'>
            <div style='display:flex;align-items:center;gap:16px;'>
                <div style='flex-shrink:0;'>{EDIS_LOGO_SVG}</div>
                <div>
                    <div style='font-size:12px;color:#2563EB;font-weight:700;letter-spacing:0.06em;
                                text-transform:uppercase;margin-bottom:6px;'>Welcome back!</div>
                    <div style='font-size:26px;font-weight:800;color:#0F172A;line-height:1.1;margin-bottom:6px;'>
                        Student Performance Overview</div>
                    <div style='color:#64748B;font-size:13.5px;'>
                        Empowering decisions through intelligent education analytics.</div>
                </div>
            </div>
            <div style='display:flex;flex-direction:column;align-items:flex-end;gap:6px;flex-shrink:0;'>
                <div style='background:#F1F5F9;border-radius:10px;padding:10px 18px;
                            font-size:13px;font-weight:600;color:#334155;
                            display:flex;align-items:center;gap:8px;
                            border:1px solid #E2E8F0;'>
                    📅 {today_str}
                </div>
                <div style='font-size:11px;color:#94A3B8;font-weight:500;'>
                    Multan University · EDIS System
                </div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    total  = len(df)
    high   = len(df[df['performance_label'] == 'High'])
    medium = len(df[df['performance_label'] == 'Medium'])
    low    = len(df[df['performance_label'] == 'Low'])

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(f"""<div class='stat-card'>
            <div class='stat-icon' style='background:#EFF6FF;'>👥</div>
            <div class='stat-value'>{total:,}</div>
            <div class='stat-label'>Total Enrolled</div>
            <div class='stat-sub'>All Departments</div>
            <div class='stat-bar'><div class='stat-bar-fill' style='background:#1D4ED8;width:100%;'></div></div>
        </div>""", unsafe_allow_html=True)
    with c2:
        st.markdown(f"""<div class='stat-card'>
            <div class='stat-icon' style='background:#DCFCE7;'>📈</div>
            <div class='stat-value' style='color:#16A34A;'>{high:,}</div>
            <div class='stat-label'>High Performers</div>
            <div class='stat-sub' style='color:#16A34A;font-weight:600;'>↑ {high/total*100:.1f}% of students</div>
            <div class='stat-bar'><div class='stat-bar-fill' style='background:#22C55E;width:{high/total*100:.0f}%;'></div></div>
        </div>""", unsafe_allow_html=True)
    with c3:
        st.markdown(f"""<div class='stat-card'>
            <div class='stat-icon' style='background:#FEF9C3;'>⚡</div>
            <div class='stat-value' style='color:#D97706;'>{medium:,}</div>
            <div class='stat-label'>Moderate Risk</div>
            <div class='stat-sub' style='color:#D97706;font-weight:600;'>{medium/total*100:.1f}% of students</div>
            <div class='stat-bar'><div class='stat-bar-fill' style='background:#F59E0B;width:{medium/total*100:.0f}%;'></div></div>
        </div>""", unsafe_allow_html=True)
    with c4:
        st.markdown(f"""<div class='stat-card'>
            <div class='stat-icon' style='background:#FEE2E2;'>🚨</div>
            <div class='stat-value' style='color:#DC2626;'>{low:,}</div>
            <div class='stat-label'>Critical Risk</div>
            <div class='stat-sub' style='color:#DC2626;font-weight:600;'>↑ {low/total*100:.1f}% of students</div>
            <div class='stat-bar'><div class='stat-bar-fill' style='background:#EF4444;width:{low/total*100:.0f}%;'></div></div>
        </div>""", unsafe_allow_html=True)

    st.markdown("<div style='height:16px;'></div>", unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        fig = go.Figure(data=[go.Pie(
            labels=['High Performers', 'Moderate Risk', 'Critical Risk'],
            values=[high, medium, low], hole=0.60,
            marker_colors=['#22C55E', '#F59E0B', '#EF4444'], textinfo='none',
            hovertemplate='<b>%{label}</b><br>%{value:,} students<br>%{percent}<extra></extra>'
        )])
        fig.add_annotation(text=f"<b>{total:,}</b>", x=0.5, y=0.55,
                           font=dict(size=22, color='#0F172A'), showarrow=False)
        fig.add_annotation(text="Total", x=0.5, y=0.42,
                           font=dict(size=13, color='#64748B'), showarrow=False)
        fig.update_layout(
            title=dict(text="<b>Performance Distribution</b>", font=dict(size=15, color='#0F172A')),
            paper_bgcolor='white', showlegend=True,
            legend=dict(orientation='v', x=0.75, y=0.5, font=dict(size=12, color='#374151')),
            margin=dict(t=50, b=20, l=20, r=20), height=300, plot_bgcolor='white'
        )
        st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False})

    with col2:
        dp = df.groupby(['department', 'performance_label']).size().reset_index(name='count')
        fig2 = go.Figure()
        for perf, clr, lbl in [('High','#22C55E','High'), ('Medium','#F59E0B','Medium'), ('Low','#EF4444','Low')]:
            d = dp[dp['performance_label'] == perf]
            fig2.add_trace(go.Bar(name=lbl, x=d['department'], y=d['count'],
                                  marker_color=clr, marker_line_width=0))
        fig2.update_layout(
            title=dict(text="<b>Department-wise Risk Breakdown</b>", font=dict(size=15, color='#0F172A')),
            paper_bgcolor='white', plot_bgcolor='white', barmode='group',
            legend=dict(font=dict(size=12, color='#374151'), orientation='h', y=1.08, x=0.5, xanchor='center'),
            margin=dict(t=60, b=20, l=20, r=20), height=300,
            xaxis=dict(showgrid=False, color='#64748B'),
            yaxis=dict(showgrid=True, gridcolor='#F1F5F9', color='#64748B', title='Students'),
        )
        st.plotly_chart(fig2, use_container_width=True, config={'displayModeBar': False})

    col3, col4 = st.columns(2)
    with col3:
        sample = df.sample(min(2000, len(df)), random_state=42)
        fig3 = px.scatter(sample, x='attendance_pct', y='mid_term', color='performance_label',
                          color_discrete_map={'High': '#22C55E', 'Medium': '#F59E0B', 'Low': '#EF4444'},
                          opacity=0.45,
                          labels={'attendance_pct': 'Attendance (%)', 'mid_term': 'Mid-Term Marks'})
        fig3.update_layout(
            title=dict(text="<b>Attendance vs Mid-Term</b>", font=dict(size=15, color='#0F172A')),
            paper_bgcolor='white', plot_bgcolor='#F8FAFC',
            legend=dict(title='', font=dict(size=12, color='#374151')),
            margin=dict(t=50, b=20, l=20, r=20), height=300,
            xaxis=dict(color='#64748B'), yaxis=dict(color='#64748B'),
        )
        st.plotly_chart(fig3, use_container_width=True, config={'displayModeBar': False})

    with col4:
        wd = df.groupby('department').apply(
            lambda x: round(len(x[x['performance_label'] == 'Low']) / len(x) * 100, 1)
        ).reset_index()
        wd.columns = ['Department', 'At-Risk %']
        fig4 = go.Figure(go.Bar(
            x=wd['Department'], y=wd['At-Risk %'],
            text=[f"{v}%" for v in wd['At-Risk %']], textposition='outside',
            marker=dict(color=wd['At-Risk %'],
                        colorscale=[[0,'#22C55E'],[0.5,'#F59E0B'],[1,'#EF4444']],
                        showscale=False, line_width=0)
        ))
        fig4.update_layout(
            title=dict(text="<b>At-Risk % by Department</b>", font=dict(size=15, color='#0F172A')),
            paper_bgcolor='white', plot_bgcolor='white',
            margin=dict(t=50, b=20, l=20, r=20), height=300,
            xaxis=dict(showgrid=False, color='#64748B'),
            yaxis=dict(showgrid=True, gridcolor='#F1F5F9', color='#64748B'),
        )
        st.plotly_chart(fig4, use_container_width=True, config={'displayModeBar': False})

    # ── Bottom row: Recent Assessments + Quick Actions ────────────────────────
    bot1, bot2 = st.columns([3, 2])
    with bot1:
        st.markdown("""
        <div style='background:white;border:1px solid #E2E8F0;border-radius:14px;
                    overflow:hidden;box-shadow:0 1px 8px rgba(15,27,53,0.05);'>
            <div style='padding:18px 22px;border-bottom:1px solid #F1F5F9;'>
                <span style='font-weight:700;font-size:15px;color:#0F172A;'>📋 Recent Assessments</span>
            </div>
            <table style='width:100%;border-collapse:collapse;'>
                <thead><tr style='background:#F8FAFC;'>
                    <th style='padding:10px 16px;text-align:left;font-size:12px;font-weight:700;
                               color:#64748B;border-bottom:1px solid #F1F5F9;'>Assessment Name</th>
                    <th style='padding:10px;text-align:left;font-size:12px;font-weight:700;
                               color:#64748B;border-bottom:1px solid #F1F5F9;'>Class / Program</th>
                    <th style='padding:10px;text-align:left;font-size:12px;font-weight:700;
                               color:#64748B;border-bottom:1px solid #F1F5F9;'>Students</th>
                    <th style='padding:10px;text-align:left;font-size:12px;font-weight:700;
                               color:#64748B;border-bottom:1px solid #F1F5F9;'>Status</th>
                </tr></thead>
                <tbody>
                    <tr style='border-bottom:1px solid #F8FAFC;'>
                        <td style='padding:12px 16px;font-size:13px;color:#1E293B;font-weight:500;'>AI Fundamentals Quiz</td>
                        <td style='padding:12px 10px;font-size:13px;color:#64748B;'>BS AI — Fall 2025</td>
                        <td style='padding:12px 10px;font-size:13px;color:#64748B;'>120</td>
                        <td style='padding:12px 10px;'><span style='background:#DCFCE7;color:#166534;
                            font-size:11.5px;font-weight:700;padding:3px 10px;border-radius:6px;'>Completed</span></td>
                    </tr>
                    <tr style='border-bottom:1px solid #F8FAFC;'>
                        <td style='padding:12px 16px;font-size:13px;color:#1E293B;font-weight:500;'>Data Structures Midterm</td>
                        <td style='padding:12px 10px;font-size:13px;color:#64748B;'>BS CS — Fall 2025</td>
                        <td style='padding:12px 10px;font-size:13px;color:#64748B;'>98</td>
                        <td style='padding:12px 10px;'><span style='background:#DCFCE7;color:#166534;
                            font-size:11.5px;font-weight:700;padding:3px 10px;border-radius:6px;'>Completed</span></td>
                    </tr>
                    <tr>
                        <td style='padding:12px 16px;font-size:13px;color:#1E293B;font-weight:500;'>Database Systems Quiz</td>
                        <td style='padding:12px 10px;font-size:13px;color:#64748B;'>BS CS — Fall 2025</td>
                        <td style='padding:12px 10px;font-size:13px;color:#64748B;'>112</td>
                        <td style='padding:12px 10px;'><span style='background:#FEF9C3;color:#713F12;
                            font-size:11.5px;font-weight:700;padding:3px 10px;border-radius:6px;'>In Progress</span></td>
                    </tr>
                </tbody>
            </table>
        </div>
        """, unsafe_allow_html=True)

    with bot2:
        # Quick Actions — labels on single line (no stacking), clean layout
        st.markdown("""
        <div style='background:white;border:1px solid #E2E8F0;border-radius:14px;
                    padding:18px 20px;box-shadow:0 1px 8px rgba(15,27,53,0.05);'>
            <div style='display:flex;align-items:center;gap:8px;margin-bottom:14px;'>
                <span style='font-size:18px;'>⚡</span>
                <span style='font-weight:700;font-size:15px;color:#0F172A;'>Quick Actions</span>
            </div>

            <div class='quick-action-card'>
                <div class='quick-action-icon' style='background:#EFF6FF;'>🔍</div>
                <div class='quick-action-text'>
                    <span class='quick-action-title'>Single Assessment</span>
                    <span class='quick-action-sub'>Assess individual student</span>
                </div>
            </div>

            <div class='quick-action-card'>
                <div class='quick-action-icon' style='background:#F0FDF4;'>👥</div>
                <div class='quick-action-text'>
                    <span class='quick-action-title'>Class Assessment</span>
                    <span class='quick-action-sub'>Assess entire class</span>
                </div>
            </div>

            <div class='quick-action-card'>
                <div class='quick-action-icon' style='background:#FFF7ED;'>📊</div>
                <div class='quick-action-text'>
                    <span class='quick-action-title'>Department Insights</span>
                    <span class='quick-action-sub'>Department analytics</span>
                </div>
            </div>

            <div class='quick-action-card' style='margin-bottom:0;'>
                <div class='quick-action-icon' style='background:#FFF1F2;'>⚠️</div>
                <div class='quick-action-text'>
                    <span class='quick-action-title'>At-Risk Registry</span>
                    <span class='quick-action-sub'>View flagged students</span>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE 2 — SINGLE ASSESSMENT
# ══════════════════════════════════════════════════════════════════════════════
elif page == "🔍  Single Assessment":
    st.markdown("""
    <div class='page-header'>
        <h2>🔍 Single Student Risk Assessment</h2>
        <p>Enter student data to generate an AI-based risk prediction and downloadable report.</p>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("<p class='section-title'>Student Information</p>", unsafe_allow_html=True)
    ci1, ci2, ci3 = st.columns(3)
    with ci1: s_id   = st.text_input("Student ID",   placeholder="e.g. 2024-CS-F-080")
    with ci2: s_name = st.text_input("Student Name", placeholder="e.g. Ali Hassan")
    with ci3: s_dept = st.selectbox("Department", ['CS','EE','ME','CE','SE','AI','BBA'])
    cs1, cs2, cs3 = st.columns(3)
    with cs1: s_subj  = st.text_input("Subject", placeholder="e.g. Data Structures")
    with cs2: s_sem   = st.selectbox("Semester", ['Spring 2026','Fall 2024','Spring 2025'])
    with cs3: s_email = st.text_input("Student Email", placeholder="student@example.com")
    st.markdown("---")

    st.markdown("<p class='section-title'>Assessment Configuration</p>", unsafe_allow_html=True)
    ca1, ca2, ca3 = st.columns(3)
    with ca1:
        nq  = st.selectbox("Number of Quizzes",          [0,1,2,3,4,5], index=2, key="nq")
        qmx = st.number_input("Max Marks per Quiz",       1, 50, 10, 1, key="qmx")
    with ca2:
        na  = st.selectbox("Number of Assignments",       [0,1,2,3,4,5], index=2, key="na")
        amx = st.number_input("Max Marks per Assignment", 1, 100, 20, 1, key="amx")
    with ca3:
        st.markdown("**Optional Components**")
        hl  = st.checkbox("Lab Work",      value=True, key="hl")
        hp  = st.checkbox("Project",       value=True, key="hp")
        hpt = st.checkbox("Participation", value=True, key="hpt")
        lab_max  = st.number_input("Lab Max",           1,100,30,1,key="lmx")  if hl  else 30
        proj_max = st.number_input("Project Max",       1,100,20,1,key="pmx")  if hp  else 20
        part_max = st.number_input("Participation Max", 1, 50,10,1,key="ptmx") if hpt else 10

    use_viva = False; use_practical = False; viva_max = 0; practical_max = 0
    if hl:
        st.markdown("<div class='lab-sub-section'><b style='color:#1D4ED8;font-size:13px;'>🔬 Lab Sub-Components (Optional)</b></div>",
                    unsafe_allow_html=True)
        lsc1, lsc2 = st.columns(2)
        with lsc1:
            use_viva = st.checkbox("Include Viva Evaluation", value=False, key="use_viva")
            if use_viva:
                viva_max = st.number_input("Viva Max Marks", 1, lab_max, min(10, lab_max), 1, key="viva_max")
        with lsc2:
            use_practical = st.checkbox("Include Practical / Lab Work (Mids so far)", value=False, key="use_practical")
            if use_practical:
                practical_max = st.number_input("Practical Max Marks (cumulative till Mids)",
                                                1, lab_max, min(20, lab_max), 1, key="practical_max")

    st.markdown("---")
    st.markdown("<p class='section-title'>Assessment Marks</p>", unsafe_allow_html=True)
    marks_data = []; q_vals = []; a_vals = []

    if nq > 0:
        st.markdown("**Quizzes**")
        qc = st.columns(nq)
        for i in range(nq):
            with qc[i]:
                v = st.number_input(f"Quiz {i+1} / {qmx}", 0.0, float(qmx),
                                    float(round(qmx*0.7,1)), 0.5, key=f"q{i}")
                q_vals.append(v)
                marks_data.append({'label': f'Quiz {i+1}', 'obtained': v, 'max': qmx})
    if na > 0:
        st.markdown("**Assignments**")
        ac = st.columns(na)
        for i in range(na):
            with ac[i]:
                v = st.number_input(f"Assignment {i+1} / {amx}", 0.0, float(amx),
                                    float(round(amx*0.7,1)), 0.5, key=f"a{i}")
                a_vals.append(v)
                marks_data.append({'label': f'Assignment {i+1}', 'obtained': v, 'max': amx})

    st.markdown("**Core Components**")
    mc1, mc2 = st.columns(2)
    with mc1:
        att = st.slider("Attendance (%)", 0.0, 100.0, 75.0, 1.0, key="att")
        if att < 75:
            st.markdown(f"""<div style='background:#FFF1F2;border:1px solid #FECDD3;border-left:3px solid #F43F5E;
                border-radius:8px;padding:8px 12px;margin-top:4px;'>
                <span style='color:#BE123C;font-size:12px;font-weight:700;'>
                {att:.0f}% attendance — Below 75% threshold. Student will be marked EXAM INELIGIBLE.</span>
            </div>""", unsafe_allow_html=True)
        mid = st.slider("Mid-Term (/ 50)", 0.0, 50.0, 35.0, 0.5, key="mid")
        marks_data += [
            {'label': 'Attendance', 'obtained': att, 'max': 100},
            {'label': 'Mid-Term',   'obtained': mid, 'max': 50}
        ]
    with mc2:
        lab  = st.slider(f"Lab Work Total (/ {lab_max})",  0.0, float(lab_max),  float(lab_max*0.7),  0.5, key="lab")  if hl  else 0.0
        proj = st.slider(f"Project (/ {proj_max})",         0.0, float(proj_max), float(proj_max*0.7), 0.5, key="proj") if hp  else 0.0
        part = st.slider(f"Participation (/ {part_max})",   0.0, float(part_max), float(part_max*0.7), 0.5, key="part") if hpt else 0.0
        if hl:  marks_data.append({'label': 'Lab Work Total', 'obtained': lab,  'max': lab_max})
        if hp:  marks_data.append({'label': 'Project',        'obtained': proj, 'max': proj_max})
        if hpt: marks_data.append({'label': 'Participation',  'obtained': part, 'max': part_max})

    if hl and use_viva:
        viva_val = st.slider(f"Viva (/ {viva_max})", 0.0, float(viva_max), float(viva_max*0.7), 0.5, key="viva_val")
        marks_data.append({'label': 'Lab Viva', 'obtained': viva_val, 'max': viva_max})
    if hl and use_practical:
        prac_val = st.slider(f"Practical / Lab Work till Mids (/ {practical_max})",
                              0.0, float(practical_max), float(practical_max*0.7), 0.5, key="practical_val")
        marks_data.append({'label': 'Lab Practical (till Mids)', 'obtained': prac_val, 'max': practical_max})

    st.markdown("---")
    if st.button("🔍  Assess Student Risk & Generate Report", use_container_width=True):
        label, risk_pct_str, overall_pct, att_ineligible, att_pct = predict_and_classify(marks_data)
        result_card(label, risk_pct_str, s_name or s_id or "Student", s_dept,
                    s_subj or "N/A", overall_pct, att_ineligible, att_pct)

        grade, next_g, pct, improvements = analyze_grade(marks_data)
        st.markdown("---")
        st.markdown("<p class='section-title'>Grade Analyser</p>", unsafe_allow_html=True)
        ga1, ga2, ga3 = st.columns(3)
        ga1.metric("Current Grade", grade)
        ga2.metric("Score", f"{pct}%")
        ga3.metric("Next Grade Target", next_g if next_g != '—' else "A+ (Achieved!)")
        if improvements:
            st.markdown("**Areas to improve for next grade:**")
            for imp in improvements[:3]:
                icon = "🔴" if imp['priority'] == 'High' else "🟡"
                st.info(f"{icon} **{imp['area']}** — currently {imp['current']}% (+{imp['potential']} marks available)")

        st.markdown("---")
        st.markdown("<p class='section-title'>Teacher Recommendation</p>", unsafe_allow_html=True)
        auto_issue      = auto_detect_issue(marks_data, label)
        auto_suggestion = ISSUE_SUGGESTIONS.get(auto_issue, "")
        st.markdown(f"""<div class='suggestion-box'>
            <p style='color:#92400E;font-weight:700;margin:0 0 4px 0;font-size:13.5px;'>
                Auto-Detected Issue: <span style='color:#B45309;'>{auto_issue}</span></p>
            <p style='color:#78350F;font-size:12px;margin:0;'>
                Auto-identified based on marks. Edit the recommendation below before generating the report.</p>
        </div>""", unsafe_allow_html=True)

        col_iss, col_sug = st.columns([1, 2])
        with col_iss:
            issue = st.selectbox("Refine Issue (Optional)", list(ISSUE_SUGGESTIONS.keys()),
                index=list(ISSUE_SUGGESTIONS.keys()).index(auto_issue) if auto_issue in ISSUE_SUGGESTIONS else 0,
                key="iss_post")
            st.caption("Auto-selected based on marks.")
        with col_sug:
            suggestion = st.text_area("Teacher Recommendation",
                value=ISSUE_SUGGESTIONS.get(issue, auto_suggestion), height=130,
                key=f"sug_post_{issue}", placeholder="Edit freely.")

        st.markdown("---")
        s_info = {
            "Student ID":    s_id   or "N/A",
            "Student Name":  s_name or "N/A",
            "Department":    s_dept,
            "Subject":       s_subj or "N/A",
            "Semester":      s_sem,
            "Attendance":    f"{att:.1f}% {'EXAM INELIGIBLE' if att_ineligible else 'Eligible'}",
            "Current Grade": f"{grade} ({pct}%)",
            "Risk Level":    f"{risk_display(label)} ({risk_pct_str})",
            "Report Date":   datetime.datetime.now().strftime("%d %B %Y"),
        }
        pdf = make_pdf(s_info, marks_data, label, risk_pct_str, suggestion, att_ineligible)
        st.download_button("📥 Download PDF Report", pdf,
                           f"EDIS_Report_{(s_id or 'Student').replace('-','_')}.pdf",
                           "application/pdf", use_container_width=True)

        st.markdown("---")
        st.markdown("<p class='section-title'>Send Report via Email</p>", unsafe_allow_html=True)
        st.markdown("""<div class='email-section'>
            <p style='color:#1E40AF;font-weight:700;margin:0 0 4px 0;font-size:14px;'>Email Report to Student</p>
            <p style='color:#1D4ED8;font-size:13px;margin:0;'>
                The PDF report will be attached and sent directly to the student.</p>
        </div>""", unsafe_allow_html=True)
        email_col1, email_col2 = st.columns([3, 1])
        with email_col1:
            final_email = st.text_input("Recipient Email Address", value=s_email,
                placeholder="student@example.com", key="final_email")
        with email_col2:
            st.markdown("<br>", unsafe_allow_html=True)
            send_clicked = st.button("Send Email", use_container_width=True, key="send_email_btn")
        if send_clicked:
            if not final_email or "@" not in final_email:
                st.warning("Please enter a valid email address.")
            else:
                with st.spinner(f"Sending report to {final_email} ..."):
                    pdf_for_email = make_pdf(s_info, marks_data, label, risk_pct_str, suggestion, att_ineligible)
                    success, err_msg = send_email_with_pdf(
                        to_email=final_email, student_name=s_name or s_id or "Student",
                        risk_label=label, subject_name=s_subj or "N/A",
                        pdf_buffer=pdf_for_email, suggestion_text=suggestion,
                    )
                if success:
                    st.success(f"Report successfully sent to {final_email}!")
                    st.balloons()
                else:
                    st.error(f"Email failed: {err_msg}")


# ══════════════════════════════════════════════════════════════════════════════
# PAGE 3 — CLASS ASSESSMENT
# Student ID + Name shown on one clear line; no black overlay block
# ══════════════════════════════════════════════════════════════════════════════
elif page == "👥  Class Assessment":
    st.markdown("""
    <div class='page-header'>
        <h2>👥 Class Assessment</h2>
        <p>Import your student list from a PDF or Excel file, configure the assessment structure,
           then fill marks for each student. All components are included in the final report.</p>
    </div>
    """, unsafe_allow_html=True)

    if 'batch'               not in st.session_state: st.session_state.batch = []
    if 'class_students'      not in st.session_state: st.session_state.class_students = []
    if 'current_student_idx' not in st.session_state: st.session_state.current_student_idx = 0

    # ── Step 1: Import ────────────────────────────────────────────────────────
    st.markdown("<p class='section-title'>Step 1 — Import Student List</p>", unsafe_allow_html=True)
    tab_pdf, tab_excel, tab_manual = st.tabs(["Upload PDF", "Upload Excel", "Add Manually"])

    with tab_pdf:
        st.markdown("""<div class='pdf-import-section'>
            <p style='color:#15803D;font-weight:700;margin:0 0 4px 0;font-size:14px;'>PDF Class List — Auto Import</p>
            <p style='color:#166534;font-size:12px;margin:0;'>
                Upload your class list PDF. Roll numbers (e.g. <b>2024-CS-F-035</b>) and names are extracted automatically.</p>
        </div>""", unsafe_allow_html=True)
        pdf_file = st.file_uploader("Choose a PDF file", type=['pdf'], key="class_pdf_uploader")
        if pdf_file is not None:
            if st.button("Extract Students from PDF", use_container_width=True, key="extract_pdf"):
                with st.spinner("Extracting student data from PDF..."):
                    extracted = parse_students_from_pdf(pdf_file)
                if extracted:
                    existing_keys = {(s['name'].lower(), s['roll_no'].lower())
                                     for s in st.session_state.class_students}
                    added = 0
                    for s in extracted:
                        key = (s['name'].lower(), s['roll_no'].lower())
                        if key not in existing_keys:
                            st.session_state.class_students.append(s)
                            existing_keys.add(key)
                            added += 1
                    st.success(f"✅ {added} student(s) imported! ({len(extracted)-added} duplicates skipped)")
                    st.rerun()
                else:
                    st.warning("Could not extract students. Try Excel or add manually.")

    with tab_excel:
        st.markdown("""<div style='background:#EFF6FF;border:1px solid #BFDBFE;border-left:4px solid #3B82F6;
            border-radius:14px;padding:18px 22px;margin:8px 0;'>
            <p style='color:#1E40AF;font-weight:700;margin:0 0 4px 0;font-size:14px;'>Excel Class List</p>
            <p style='color:#1D4ED8;font-size:12px;margin:0;'>
                Upload an Excel (.xlsx) with columns: <b>Name</b> and <b>Roll No / ID</b>.</p>
        </div>""", unsafe_allow_html=True)
        excel_file = st.file_uploader("Choose an Excel file (.xlsx)", type=['xlsx','xls'], key="class_excel_uploader")
        if excel_file is not None:
            if st.button("Extract Students from Excel", use_container_width=True, key="extract_excel"):
                with st.spinner("Extracting from Excel..."):
                    extracted = parse_students_from_excel(excel_file)
                if extracted:
                    existing_keys = {(s['name'].lower(), s['roll_no'].lower())
                                     for s in st.session_state.class_students}
                    added = 0
                    for s in extracted:
                        key = (s['name'].lower(), s['roll_no'].lower())
                        if key not in existing_keys:
                            st.session_state.class_students.append(s)
                            existing_keys.add(key)
                            added += 1
                    st.success(f"✅ {added} student(s) imported from Excel!")
                    st.rerun()
                else:
                    st.warning("Could not extract students. Check columns.")

    with tab_manual:
        st.markdown("**Add Student Manually**")
        man_col1, man_col2, man_col3 = st.columns([2, 2, 1])
        with man_col1: man_name = st.text_input("Student Full Name",        placeholder="e.g. Muhammad Ali Hassan", key="man_name")
        with man_col2: man_roll = st.text_input("Roll Number / Student ID", placeholder="e.g. 2024-CS-F-080",      key="man_roll")
        with man_col3:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("Add to List", use_container_width=True, key="man_add_list"):
                if man_name.strip():
                    st.session_state.class_students.append(
                        {'name': man_name.strip(), 'roll_no': man_roll.strip() or 'N/A'})
                    st.success(f"✅ {man_name} added.")
                    st.rerun()
                else:
                    st.warning("Please enter at least a student name.")

    if st.session_state.class_students:
        list_col1, list_col2 = st.columns([4, 1])
        with list_col1:
            st.markdown(f"**Student List — {len(st.session_state.class_students)} students loaded**")
        with list_col2:
            if st.button("Clear List", use_container_width=True, key="clear_students"):
                st.session_state.class_students = []
                st.session_state.current_student_idx = 0
                st.rerun()
        preview_rows = [{"#": i+1, "Roll No / ID": s.get('roll_no','N/A'), "Full Name": s.get('name','')}
                        for i, s in enumerate(st.session_state.class_students)]
        st.dataframe(pd.DataFrame(preview_rows), use_container_width=True,
                     height=min(220, 40+len(preview_rows)*36), hide_index=True)

    st.markdown("---")

    # ── Step 2: Configuration ─────────────────────────────────────────────────
    st.markdown("<p class='section-title'>Step 2 — Assessment Configuration</p>", unsafe_allow_html=True)
    cc1, cc2, cc3 = st.columns(3)
    with cc1:
        c_nq  = st.selectbox("Number of Quizzes",          [0,1,2,3,4,5], index=2, key="c_nq")
        c_qmx = st.number_input("Max Marks per Quiz",       1, 50, 10, 1,           key="c_qmx")
    with cc2:
        c_na  = st.selectbox("Number of Assignments",       [0,1,2,3,4,5], index=2, key="c_na")
        c_amx = st.number_input("Max Marks per Assignment", 1, 100, 20, 1,           key="c_amx")
    with cc3:
        st.markdown("**Optional Components**")
        c_hl  = st.checkbox("Lab Work",      value=True, key="c_hl")
        c_hp  = st.checkbox("Project",       value=True, key="c_hp")
        c_hpt = st.checkbox("Participation", value=True, key="c_hpt")
        c_lmx  = st.number_input("Lab Max",           1,100,30,1,key="c_lmx")  if c_hl  else 30
        c_pmx  = st.number_input("Project Max",       1,100,20,1,key="c_pmx")  if c_hp  else 20
        c_ptmx = st.number_input("Participation Max", 1, 50,10,1,key="c_ptmx") if c_hpt else 10

    c_use_viva = False; c_use_practical = False; c_viva_max = 0; c_practical_max = 0
    if c_hl:
        st.markdown("<div class='lab-sub-section'><b style='color:#1D4ED8;font-size:13px;'>🔬 Lab Sub-Components (Optional)</b></div>",
                    unsafe_allow_html=True)
        clsc1, clsc2 = st.columns(2)
        with clsc1:
            c_use_viva = st.checkbox("Include Viva Evaluation", value=False, key="c_use_viva")
            if c_use_viva:
                c_viva_max = st.number_input("Viva Max Marks", 1, c_lmx, min(10, c_lmx), 1, key="c_viva_max")
        with clsc2:
            c_use_practical = st.checkbox("Include Practical / Lab Work (till Mids)", value=False, key="c_use_practical")
            if c_use_practical:
                c_practical_max = st.number_input("Practical Max Marks",
                                                   1, c_lmx, min(20, c_lmx), 1, key="c_practical_max")

    st.markdown("---")

    # ── Step 3: Class info ────────────────────────────────────────────────────
    st.markdown("<p class='section-title'>Step 3 — Class / Subject Information</p>", unsafe_allow_html=True)
    ci1, ci2, ci3, ci4 = st.columns(4)
    with ci1: c_dept    = st.selectbox("Department", ['CS','EE','ME','CE','SE','AI','BBA'], key="c_dept_sel")
    with ci2: c_subj    = st.text_input("Subject Name",    placeholder="e.g. Data Structures", key="c_subj_name")
    with ci3: c_sem     = st.selectbox("Semester", ['Fall 2025','Spring 2026','Fall 2024','Spring 2025'], key="c_sem_sel")
    with ci4: c_section = st.text_input("Section / Group", placeholder="e.g. Section A", key="c_section")

    st.markdown("---")

    # ── Step 4: Enter marks ───────────────────────────────────────────────────
    st.markdown("<p class='section-title'>Step 4 — Enter Marks for Each Student</p>", unsafe_allow_html=True)

    if not st.session_state.class_students:
        st.info("No students loaded. Import a PDF/Excel or add manually in Step 1.")
    else:
        total_students   = len(st.session_state.class_students)
        already_assessed = [r['Student Name'] for r in st.session_state.batch]
        assessed_count   = len(st.session_state.batch)
        progress_pct     = int(assessed_count / total_students * 100) if total_students > 0 else 0

        st.markdown(f"""
        <div style='background:white;border:1px solid #E2E8F0;border-radius:12px;
                    padding:16px 20px;margin-bottom:16px;box-shadow:0 1px 6px rgba(15,27,53,0.05);'>
            <div style='display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;'>
                <span style='font-weight:700;color:#374151;font-size:14px;'>Assessment Progress</span>
                <span style='font-weight:800;color:#1D4ED8;font-size:15px;'>{assessed_count} / {total_students} Students</span>
            </div>
            <div style='background:#F1F5F9;border-radius:8px;height:8px;'>
                <div style='background:linear-gradient(90deg,#1D4ED8,#3B82F6);border-radius:8px;
                            height:8px;width:{progress_pct}%;transition:width 0.3s;'></div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        # Dropdown: "ROLL_NO — Full Name" format
        student_options = [f"{s['roll_no']} — {s['name']}" for s in st.session_state.class_students]
        sel_idx = st.selectbox("Select Student to Assess", range(len(student_options)),
            format_func=lambda i: (
                f"✅ {student_options[i]}"
                if st.session_state.class_students[i]['name'] in already_assessed
                else f"⏳ {student_options[i]}"
            ), key="student_selector")

        cur_student  = st.session_state.class_students[sel_idx]
        already_done = cur_student['name'] in already_assessed

        # ── Student card — clean HTML, no dark block, ID + Name on one line ──
        assessed_badge = (
            '<div style="color:#16A34A;font-size:11.5px;margin-top:4px;font-weight:600;">✅ Already assessed</div>'
            if already_done else ''
        )
        section_line = (f' &nbsp;·&nbsp; {c_section}') if c_section else ''

        st.markdown(f"""
        <div class='student-card'>
            <div style='display:flex;align-items:center;gap:14px;'>
                <div style='background:linear-gradient(135deg,#1D4ED8,#3B82F6);color:white;
                            border-radius:50%;width:46px;height:46px;display:flex;
                            align-items:center;justify-content:center;font-weight:800;
                            font-size:18px;flex-shrink:0;'>{cur_student['name'][0].upper()}</div>
                <div style='flex:1;min-width:0;'>
                    <div style='display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin-bottom:6px;'>
                        <span style='font-weight:800;color:#0F172A;font-size:16px;'>
                            {cur_student['name']}</span>
                        <span style='background:#EFF6FF;color:#1D4ED8;font-size:12px;font-weight:700;
                                     padding:3px 10px;border-radius:20px;border:1px solid #BFDBFE;
                                     white-space:nowrap;'>
                            {cur_student['roll_no']}</span>
                    </div>
                    <div style='display:flex;flex-wrap:wrap;gap:6px;align-items:center;'>
                        <span style='color:#64748B;font-size:12.5px;font-weight:500;'>{c_dept}</span>
                        <span style='color:#CBD5E1;font-size:12px;'>·</span>
                        <span style='color:#64748B;font-size:12.5px;font-weight:500;'>{c_subj or 'N/A'}</span>
                        <span style='color:#CBD5E1;font-size:12px;'>·</span>
                        <span style='color:#64748B;font-size:12.5px;font-weight:500;'>{c_sem}{section_line}</span>
                    </div>
                    {assessed_badge}
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        beml = st.text_input("Student Email (optional)",
                              placeholder="student@university.edu", key=f"beml_{sel_idx}")

        bq_vals = []; ba_vals = []; b_marks_data = []
        if c_nq > 0:
            st.markdown("**Quizzes**")
            bqc = st.columns(c_nq)
            for i in range(c_nq):
                with bqc[i]:
                    v = st.number_input(f"Quiz {i+1} / {c_qmx}", 0.0, float(c_qmx),
                                        float(round(c_qmx*0.7,1)), 0.5, key=f"bq{i}_{sel_idx}")
                    bq_vals.append(v)
                    b_marks_data.append({'label': f'Quiz {i+1}', 'obtained': v, 'max': c_qmx})
        if c_na > 0:
            st.markdown("**Assignments**")
            bac = st.columns(c_na)
            for i in range(c_na):
                with bac[i]:
                    v = st.number_input(f"Assignment {i+1} / {c_amx}", 0.0, float(c_amx),
                                        float(round(c_amx*0.7,1)), 0.5, key=f"ba{i}_{sel_idx}")
                    ba_vals.append(v)
                    b_marks_data.append({'label': f'Assignment {i+1}', 'obtained': v, 'max': c_amx})

        st.markdown("**Core Components**")
        bcore1, bcore2 = st.columns(2)
        with bcore1:
            batt = st.slider("Attendance (%)", 0.0, 100.0, 75.0, 1.0, key=f"batt_{sel_idx}")
            if batt < 75:
                st.markdown(f"""<div style='background:#FFF1F2;border:1px solid #FECDD3;border-left:3px solid #F43F5E;
                    border-radius:8px;padding:8px 12px;margin-top:4px;'>
                    <span style='color:#BE123C;font-size:12px;font-weight:700;'>
                    {batt:.0f}% — EXAM INELIGIBLE (below 75%)</span>
                </div>""", unsafe_allow_html=True)
            bmid = st.slider("Mid-Term (/ 50)", 0.0, 50.0, 35.0, 0.5, key=f"bmid_{sel_idx}")
            b_marks_data += [
                {'label': 'Attendance', 'obtained': batt, 'max': 100},
                {'label': 'Mid-Term',   'obtained': bmid, 'max': 50}
            ]
        with bcore2:
            blab  = st.slider(f"Lab Work Total (/ {c_lmx})",  0.0, float(c_lmx),  float(c_lmx*0.7),  0.5, key=f"blab_{sel_idx}")  if c_hl  else 0.0
            bproj = st.slider(f"Project (/ {c_pmx})",          0.0, float(c_pmx),  float(c_pmx*0.7),  0.5, key=f"bproj_{sel_idx}") if c_hp  else 0.0
            bpart = st.slider(f"Participation (/ {c_ptmx})",   0.0, float(c_ptmx), float(c_ptmx*0.7), 0.5, key=f"bpart_{sel_idx}") if c_hpt else 0.0
            if c_hl:  b_marks_data.append({'label': 'Lab Work Total', 'obtained': blab,  'max': c_lmx})
            if c_hp:  b_marks_data.append({'label': 'Project',        'obtained': bproj, 'max': c_pmx})
            if c_hpt: b_marks_data.append({'label': 'Participation',  'obtained': bpart, 'max': c_ptmx})

        b_viva_val = 0.0; b_practical_val = 0.0
        if c_hl and c_use_viva:
            b_viva_val = st.slider(f"Viva (/ {c_viva_max})", 0.0, float(c_viva_max),
                                    float(c_viva_max*0.7), 0.5, key=f"bviva_{sel_idx}")
            b_marks_data.append({'label': 'Lab Viva', 'obtained': b_viva_val, 'max': c_viva_max})
        if c_hl and c_use_practical:
            b_practical_val = st.slider(f"Practical / Lab Work till Mids (/ {c_practical_max})",
                                         0.0, float(c_practical_max), float(c_practical_max*0.7),
                                         0.5, key=f"bprac_{sel_idx}")
            b_marks_data.append({'label': 'Lab Practical (till Mids)', 'obtained': b_practical_val, 'max': c_practical_max})

        if st.button(f"➕  Add {cur_student['name']} to Class Record", use_container_width=True, key="badd"):
            b_label, b_risk_pct_str, b_overall_pct, b_att_ineligible, b_att_pct = predict_and_classify(b_marks_data)
            rd          = risk_display(b_label)
            exam_status = "INELIGIBLE" if b_att_ineligible else "Eligible"
            row = {
                "Student ID":   cur_student['roll_no'],
                "Student Name": cur_student['name'],
                "Department":   c_dept,
                "Subject":      c_subj or "N/A",
                "Section":      c_section or "N/A",
                "Semester":     c_sem,
                "Email":        beml or "N/A",
                "Attendance %": batt,
                "Exam Status":  exam_status,
                "Mid-Term /50": bmid,
                "Overall %":    round(b_overall_pct, 1),
                "Risk %":       b_risk_pct_str,
                "Risk Level":   rd,
            }
            for i, v in enumerate(bq_vals):   row[f"Quiz {i+1} /{c_qmx}"]     = v
            for i, v in enumerate(ba_vals):   row[f"Assign {i+1} /{c_amx}"]   = v
            if c_hl:            row[f"Lab /{c_lmx}"]           = blab
            if c_hp:            row[f"Project /{c_pmx}"]       = bproj
            if c_hpt:           row[f"Particip /{c_ptmx}"]     = bpart
            if c_hl and c_use_viva:      row[f"Viva /{c_viva_max}"]      = b_viva_val
            if c_hl and c_use_practical: row[f"Prac /{c_practical_max}"] = b_practical_val

            st.session_state.batch.append(row)
            msg = f"✅ {cur_student['name']} added — **{rd}** ({b_overall_pct:.1f}% overall)"
            if b_att_ineligible:
                msg += f" · EXAM INELIGIBLE (Attendance {b_att_pct:.1f}%)"
            st.success(msg)
            assessed_names = [r['Student Name'] for r in st.session_state.batch]
            for i, s in enumerate(st.session_state.class_students):
                if s['name'] not in assessed_names:
                    st.session_state.current_student_idx = i
                    break
            st.rerun()

    # ── Class Record ──────────────────────────────────────────────────────────
    if st.session_state.batch:
        st.markdown("---")
        bdf = pd.DataFrame(st.session_state.batch)
        st.markdown(f"<p class='section-title'>Class Record — {len(bdf)} Students Assessed</p>",
                    unsafe_allow_html=True)

        def cr(v):
            if v == 'High Risk':       return 'background-color:#FEE2E2;color:#991B1B;font-weight:700'
            elif v == 'Moderate Risk': return 'background-color:#FEF9C3;color:#713F12;font-weight:700'
            return 'background-color:#DCFCE7;color:#166534;font-weight:700'

        def ce(v):
            if v == 'INELIGIBLE': return 'background-color:#FFF1F2;color:#BE123C;font-weight:700'
            return 'background-color:#F0FDF4;color:#166534;font-weight:600'

        show_cols = [c for c in bdf.columns if c != 'Email']
        styled = bdf[show_cols].style.map(cr, subset=['Risk Level'])
        if 'Exam Status' in show_cols:
            styled = styled.map(ce, subset=['Exam Status'])
        st.dataframe(styled, use_container_width=True, height=320)

        r1, r2, r3 = st.columns(3)
        r1.metric("High Risk",     len(bdf[bdf['Risk Level'] == 'High Risk']))
        r2.metric("Moderate Risk", len(bdf[bdf['Risk Level'] == 'Moderate Risk']))
        r3.metric("Low Risk",      len(bdf[bdf['Risk Level'] == 'Low Risk']))
        av1, av2, av3 = st.columns(3)
        av1.metric("Avg Attendance", f"{bdf['Attendance %'].mean():.1f}%")
        av2.metric("Avg Mid-Term",   f"{bdf['Mid-Term /50'].mean():.1f}/50")
        if 'Exam Status' in bdf.columns:
            inelig_count = len(bdf[bdf['Exam Status'] == 'INELIGIBLE'])
            av3.metric("Exam Ineligible", str(inelig_count),
                       delta=f"{inelig_count/len(bdf)*100:.1f}%" if len(bdf) > 0 else "0%",
                       delta_color="inverse")

        st.markdown("---")
        st.markdown("<p class='section-title'>Bulk Email — At-Risk Students</p>", unsafe_allow_html=True)
        st.markdown("""<div class='email-section'>
            <p style='color:#1E40AF;font-weight:700;margin:0 0 4px 0;font-size:14px;'>
                Send Notifications to All High-Risk Students</p>
            <p style='color:#1D4ED8;font-size:13px;margin:0;'>
                A warning email will be sent to every High-Risk student with a valid email address.</p>
        </div>""", unsafe_allow_html=True)

        high_risk_with_email = bdf[
            (bdf['Risk Level'] == 'High Risk') &
            (bdf['Email'].notna()) & (bdf['Email'] != 'N/A') &
            (bdf['Email'].str.contains('@', na=False))
        ]
        st.info(f"**{len(high_risk_with_email)}** high-risk student(s) with valid email found.")

        if st.button("Send Bulk Emails to High-Risk Students", use_container_width=True, key="bulk_email"):
            if len(high_risk_with_email) == 0:
                st.warning("No high-risk students with valid email addresses found.")
            else:
                sent = 0; failed = []
                prog = st.progress(0, text="Sending emails...")
                for idx, (_, row) in enumerate(high_risk_with_email.iterrows()):
                    is_inelig = row.get('Exam Status', '') == 'INELIGIBLE'
                    suggestion_bulk = (
                        "URGENT: Your attendance is below 75% and you are NOT ELIGIBLE for the final exam. "
                        "Please contact your course teacher immediately."
                        if is_inelig else
                        "You are classified as HIGH RISK. Please meet your teacher and advisor immediately."
                    )
                    bulk_marks = [
                        {'label': 'Attendance', 'obtained': row.get('Attendance %', 0), 'max': 100},
                        {'label': 'Mid-Term',   'obtained': row.get('Mid-Term /50', 0), 'max': 50},
                    ]
                    for q in range(1, 6):
                        k = f"Quiz {q} /{c_qmx}"
                        if k in row and row[k] not in [None, 'N/A']:
                            bulk_marks.append({'label': f'Quiz {q}', 'obtained': float(row[k]), 'max': c_qmx})
                    for a in range(1, 6):
                        k = f"Assign {a} /{c_amx}"
                        if k in row and row[k] not in [None, 'N/A']:
                            bulk_marks.append({'label': f'Assignment {a}', 'obtained': float(row[k]), 'max': c_amx})
                    if c_hl:
                        k = f"Lab /{c_lmx}"
                        if k in row and row[k] not in [None, 'N/A']:
                            bulk_marks.append({'label': 'Lab Work', 'obtained': float(row[k]), 'max': c_lmx})

                    risk_pct_val = str(row.get('Risk %', '—'))
                    bulk_info = {
                        "Student ID":   row.get('Student ID',   'N/A'),
                        "Student Name": row.get('Student Name', 'N/A'),
                        "Department":   row.get('Department',   'N/A'),
                        "Subject":      row.get('Subject',      'N/A'),
                        "Section":      row.get('Section',      'N/A'),
                        "Semester":     row.get('Semester',     'N/A'),
                        "Attendance":   f"{row.get('Attendance %',0):.1f}% {'EXAM INELIGIBLE' if is_inelig else ''}",
                        "Risk Level":   f"High Risk ({risk_pct_val})",
                        "Report Date":  datetime.datetime.now().strftime("%d %B %Y"),
                    }
                    bulk_pdf = make_pdf(bulk_info, bulk_marks, 'Low', risk_pct_val, suggestion_bulk, is_inelig)
                    ok, err = send_email_with_pdf(
                        to_email=row['Email'], student_name=row.get('Student Name', 'Student'),
                        risk_label='Low', subject_name=row.get('Subject', 'N/A'),
                        pdf_buffer=bulk_pdf, suggestion_text=suggestion_bulk,
                    )
                    if ok: sent += 1
                    else:  failed.append(f"{row.get('Student Name','?')} ({row['Email']}): {err}")
                    prog.progress((idx+1)/len(high_risk_with_email),
                                  text=f"Sending {idx+1}/{len(high_risk_with_email)}...")
                prog.empty()
                if sent > 0:  st.success(f"✅ {sent} email(s) sent successfully!")
                if failed:
                    st.error(f"❌ {len(failed)} email(s) failed:")
                    for f_msg in failed: st.caption(f"• {f_msg}")

        st.markdown("---")
        st.markdown("<p class='section-title'>Export Class Report</p>", unsafe_allow_html=True)
        d1, d2, d3, d4 = st.columns(4)
        with d1:
            st.download_button("Download CSV", bdf.to_csv(index=False),
                               "EDIS_class_results.csv", "text/csv", use_container_width=True)
        with d2:
            xl = make_excel(bdf)
            st.download_button("Download Excel (.xlsx)", xl, "EDIS_class_results.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               use_container_width=True)
        with d3:
            class_info = {'department': c_dept, 'semester': c_sem,
                          'subject': c_subj or 'N/A', 'section': c_section or 'N/A'}
            class_pdf = make_class_pdf(bdf, class_info)
            st.download_button("Download Class PDF", class_pdf, "EDIS_class_report.pdf",
                               "application/pdf", use_container_width=True)
        with d4:
            if st.button("Clear All Records", use_container_width=True, key="clr"):
                st.session_state.batch = []
                st.rerun()
    else:
        if not st.session_state.class_students:
            st.info("Import your class list in Step 1 to get started.")
        else:
            st.info(f"{len(st.session_state.class_students)} student(s) loaded. "
                    "Select a student and click Add to Class Record after entering marks.")


# ══════════════════════════════════════════════════════════════════════════════
# PAGE 4 — DEPARTMENT INSIGHTS
# ══════════════════════════════════════════════════════════════════════════════
elif page == "📊  Department Insights":
    st.markdown("""
    <div class='page-header'>
        <h2>📊 Department-wise Academic Insights</h2>
        <p>Drill down into performance patterns across departments and semesters.</p>
    </div>
    """, unsafe_allow_html=True)

    dept_f = st.selectbox("Department", ['All Departments'] + sorted(df['department'].unique().tolist()))
    filt   = df if dept_f == 'All Departments' else df[df['department'] == dept_f]

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Students", f"{len(filt):,}")
    c2.metric("Avg Attendance", f"{filt['attendance_pct'].mean():.1f}%")
    c3.metric("Avg Mid-Term",   f"{filt['mid_term'].mean():.1f} / 50")
    c4.metric("At-Risk Count",  f"{len(filt[filt['performance_label']=='Low']):,}", delta_color="inverse")

    st.markdown("---")
    col1, col2 = st.columns(2)
    with col1:
        wd = df.groupby('department').apply(
            lambda x: round(len(x[x['performance_label'] == 'Low']) / len(x) * 100, 1)
        ).reset_index()
        wd.columns = ['Department', 'At-Risk %']
        fig = go.Figure(go.Bar(
            x=wd['Department'], y=wd['At-Risk %'],
            text=[f"{v}%" for v in wd['At-Risk %']], textposition='outside',
            marker=dict(color=wd['At-Risk %'],
                        colorscale=[[0,'#22C55E'],[0.5,'#F59E0B'],[1,'#EF4444']],
                        showscale=False, line_width=0)
        ))
        fig.update_layout(
            title=dict(text="<b>At-Risk % by Department</b>", font=dict(size=15, color='#0F172A')),
            paper_bgcolor='white', plot_bgcolor='white',
            margin=dict(t=50, b=20, l=20, r=20), height=320,
            xaxis=dict(showgrid=False, color='#64748B'),
            yaxis=dict(showgrid=True, gridcolor='#F1F5F9', color='#64748B')
        )
        st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False})

    with col2:
        sp = filt.groupby(['semester', 'performance_label']).size().reset_index(name='count')
        fig2 = go.Figure()
        for perf, clr, lbl in [('High','#22C55E','High'), ('Medium','#F59E0B','Medium'), ('Low','#EF4444','Low')]:
            d = sp[sp['performance_label'] == perf]
            fig2.add_trace(go.Bar(name=lbl, x=d['semester'], y=d['count'],
                                  marker_color=clr, marker_line_width=0))
        fig2.update_layout(
            title=dict(text="<b>Semester-wise Performance</b>", font=dict(size=15, color='#0F172A')),
            paper_bgcolor='white', plot_bgcolor='white', barmode='group',
            legend=dict(font=dict(size=12, color='#374151'), orientation='h', y=1.08, x=0.5, xanchor='center'),
            margin=dict(t=60, b=20, l=20, r=20), height=320,
            xaxis=dict(showgrid=False, color='#64748B'),
            yaxis=dict(showgrid=True, gridcolor='#F1F5F9', color='#64748B')
        )
        st.plotly_chart(fig2, use_container_width=True, config={'displayModeBar': False})

    sa = filt.groupby('subject')['mid_term'].mean().sort_values().reset_index()
    fig3 = go.Figure(go.Bar(
        x=sa['mid_term'], y=sa['subject'], orientation='h',
        text=[f"{v:.1f}" for v in sa['mid_term']], textposition='outside',
        marker=dict(color=sa['mid_term'],
                    colorscale=[[0,'#BFDBFE'],[1,'#1D4ED8']], showscale=False, line_width=0)
    ))
    fig3.update_layout(
        title=dict(text="<b>Subject-wise Avg Mid-Term</b>", font=dict(size=15, color='#0F172A')),
        paper_bgcolor='white', plot_bgcolor='white',
        margin=dict(t=50, b=20, l=20, r=20), height=420,
        xaxis=dict(color='#64748B'), yaxis=dict(color='#64748B', title='')
    )
    st.plotly_chart(fig3, use_container_width=True, config={'displayModeBar': False})


# ══════════════════════════════════════════════════════════════════════════════
# PAGE 5 — AT-RISK REGISTRY
# ══════════════════════════════════════════════════════════════════════════════
elif page == "⚠️  At-Risk Registry":
    st.markdown("""
    <div class='page-header'>
        <h2>⚠️ At-Risk Student Registry</h2>
        <p>Students flagged as high-risk based on current performance. Immediate academic support recommended.</p>
    </div>
    """, unsafe_allow_html=True)

    f1, f2 = st.columns(2)
    with f1: dpt = st.selectbox("Department", ['All Departments'] + sorted(df['department'].unique().tolist()))
    with f2: sem = st.selectbox("Semester",   ['All Semesters']   + sorted(df['semester'].unique().tolist()))

    ar = df[df['performance_label'] == 'Low'].copy()
    if dpt != 'All Departments': ar = ar[ar['department'] == dpt]
    if sem != 'All Semesters':   ar = ar[ar['semester']   == sem]

    if 'attendance_pct' in ar.columns:
        inelig_count = len(ar[ar['attendance_pct'] < 75])
        if inelig_count > 0:
            st.markdown(f"""<div class='attendance-warning'>
                <div style='display:flex;align-items:center;gap:10px;'>
                    <span style='font-size:24px;'>🚫</span>
                    <div>
                        <div style='font-size:15px;font-weight:800;color:#BE123C;'>
                            {inelig_count} student(s) with attendance below 75% — NOT ELIGIBLE for final exam</div>
                        <div style='font-size:12.5px;color:#E11D48;margin-top:2px;'>
                            These students must be contacted immediately.</div>
                    </div>
                </div>
            </div>""", unsafe_allow_html=True)

    st.error(f"🚨 {len(ar):,} students require immediate academic attention.")

    cols = ['student_id','department','semester','subject',
            'attendance_pct','mid_term','quiz1','quiz2','quiz3',
            'assignment1','assignment2','assignment3']
    available_cols = [c for c in cols if c in ar.columns]
    display_ar = ar[available_cols].sort_values('mid_term').head(200).reset_index(drop=True)

    def style_att(v):
        if isinstance(v, (int, float)) and v < 75:
            return 'background-color:#FFF1F2;color:#BE123C;font-weight:700'
        return ''

    styled_ar = display_ar.style.map(style_att, subset=['attendance_pct'])
    st.dataframe(styled_ar, use_container_width=True, height=420)

    e1, e2 = st.columns(2)
    with e1:
        st.download_button("Export CSV", ar[available_cols].to_csv(index=False),
                           f"EDIS_AtRisk_{dpt}.csv", "text/csv", use_container_width=True)
    with e2:
        arc = ar[available_cols].copy()
        arc['Risk Level'] = 'High Risk'
        xl  = make_excel(arc)
        st.download_button("Export Excel", xl, f"EDIS_AtRisk_{dpt}.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)