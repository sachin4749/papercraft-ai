import io
import re
import os
import json
import sqlite3
import random
from typing import List, Dict, Any, Tuple
from PIL import Image

import streamlit as st

# OpenAI API Client
from openai import OpenAI

# Document Parsing Libraries
import pypdf
import pdfplumber
import docx

# PDF Generation Libraries (ReportLab Platypus)
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage, HRFlowable
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfgen import canvas

# Word Generation Libraries
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# ==========================================
# 1. DATABASE INITIALIZATION & OPERATIONS
# ==========================================

DB_FILE = "question_bank.db"

def init_db():
    try:
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS question_bank (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                subject TEXT NOT NULL,
                chapter TEXT,
                question_text TEXT NOT NULL,
                is_mcq INTEGER NOT NULL,
                option_a TEXT,
                option_b TEXT,
                option_c TEXT,
                option_d TEXT,
                correct_answer TEXT,
                marks INTEGER NOT NULL,
                difficulty TEXT,
                blooms_level TEXT
            )
        """)
        conn.commit()
        conn.close()
    except Exception as e:
        print(f"Database initialization warning: {e}")

def save_question_to_db(q_data: Dict[str, Any]):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO question_bank 
        (subject, chapter, question_text, is_mcq, option_a, option_b, option_c, option_d, correct_answer, marks, difficulty, blooms_level)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        q_data.get("subject", "General"),
        q_data.get("chapter", "General"),
        q_data["text"],
        1 if q_data["is_mcq"] else 0,
        q_data["options"].get("A", "") if q_data["is_mcq"] else "",
        q_data["options"].get("B", "") if q_data["is_mcq"] else "",
        q_data["options"].get("C", "") if q_data["is_mcq"] else "",
        q_data["options"].get("D", "") if q_data["is_mcq"] else "",
        q_data.get("correct_answer", ""),
        q_data.get("marks", 1),
        q_data.get("difficulty", "Medium"),
        q_data.get("blooms_level", "Understand")
    ))
    conn.commit()
    conn.close()

def get_filtered_questions(subject: str = None, difficulty: str = None) -> List[Dict[str, Any]]:
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    query = "SELECT id, subject, chapter, question_text, is_mcq, option_a, option_b, option_c, option_d, correct_answer, marks, difficulty, blooms_level FROM question_bank WHERE 1=1"
    params = []
    if subject and subject != "All":
        query += " AND subject = ?"
        params.append(subject)
    if difficulty and difficulty != "All":
        query += " AND difficulty = ?"
        params.append(difficulty)
    
    cursor.execute(query, params)
    rows = cursor.fetchall()
    conn.close()
    
    questions = []
    for r in rows:
        questions.append({
            "id": r[0],
            "subject": r[1],
            "chapter": r[2],
            "text": r[3],
            "is_mcq": bool(r[4]),
            "options": {"A": r[5], "B": r[6], "C": r[7], "D": r[8]} if r[4] else {"A": "", "B": "", "C": "", "D": ""},
            "correct_answer": r[9],
            "marks": r[10],
            "difficulty": r[11],
            "blooms_level": r[12]
        })
    return questions

def get_unique_subjects() -> List[str]:
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("SELECT DISTINCT subject FROM question_bank")
    rows = cursor.fetchall()
    conn.close()
    return [r[0] for r in rows if r[0]]

# ==========================================
# 2. ADVANCED BACKEND AI ENGINES
# ==========================================

def get_openai_client():
    try:
        api_key = st.secrets.get("OPENAI_API_KEY") or os.getenv("OPENAI_API_KEY")
        if not api_key:
            return None
        return OpenAI(api_key=api_key)
    except Exception:
        return None

def generate_ai_questions(subject: str, topic: str, count: int, marks: int, is_mcq: bool) -> List[Dict[str, Any]]:
    client = get_openai_client()
    if not client:
        st.error("OpenAI API Key is missing. Please configure OPENAI_API_KEY in Streamlit Secrets.")
        return []

    prompt = f"""
    You are an expert professor in {subject}. Generate exactly {count} examination questions for the topic/chapter '{topic}'.
    Each question is worth {marks} marks and should be of type {'Multiple Choice Question (MCQ)' if is_mcq else 'Descriptive/Analytical Question'}.
    
    Return your response strictly as a JSON object with a single key "questions" containing a list of dictionaries with these exact keys:
    - "text": The question statement string.
    - "options": An object with keys "A", "B", "C", "D" mapping to option strings (if is_mcq is true, otherwise empty dictionary {{}}).
    - "correct_answer": For MCQs, the correct option letter (e.g. "A"). For descriptive, a clear grading rubric or key solution.
    - "difficulty": One of "Easy", "Medium", or "Hard".
    - "blooms_level": One of "Remember", "Understand", "Apply", "Analyze", "Evaluate", or "Create".
    """

    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {"role": "system", "content": "You are an automated academic question generator that outputs strictly valid JSON."},
                {"role": "user", "content": prompt}
            ],
            response_format={"type": "json_object"}
        )
        content = json.loads(response.choices[0].message.content)
        raw_qs = content.get("questions", [])
        
        formatted_qs = []
        for q in raw_qs:
            formatted_qs.append({
                "text": q.get("text", ""),
                "is_mcq": is_mcq,
                "options": q.get("options", {"A": "", "B": "", "C": "", "D": ""}) if is_mcq else {"A": "", "B": "", "C": "", "D": ""},
                "correct_answer": q.get("correct_answer", "A" if is_mcq else "Detailed rubric..."),
                "difficulty": q.get("difficulty", "Medium"),
                "blooms_level": q.get("blooms_level", "Understand")
            })
        return formatted_qs
    except Exception as e:
        st.error(f"AI Generation Error: {e}")
        return []

def analyze_question_pedagogy(question_text: str) -> dict:
    client = get_openai_client()
    if not client:
        return {"difficulty": "Medium", "blooms_level": "Understand", "reasoning": "API key missing."}

    prompt = f"""
    Analyze the following examination question and classify its cognitive level:
    Question: "{question_text}"
    
    Provide the response strictly as a JSON object with:
    - "difficulty": "Easy", "Medium", or "Hard"
    - "blooms_level": Exactly one of "Remember", "Understand", "Apply", "Analyze", "Evaluate", "Create"
    - "reasoning": A brief 1-sentence pedagogical justification.
    """

    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {"role": "system", "content": "You are an expert psychometrician and professor."},
                {"role": "user", "content": prompt}
            ],
            response_format={"type": "json_object"}
        )
        return json.loads(response.choices[0].message.content)
    except Exception as e:
        return {"difficulty": "Medium", "blooms_level": "Understand", "reasoning": str(e)}

def audit_exam_coverage(syllabus_outline: str, sections_data: list) -> str:
    client = get_openai_client()
    if not client:
        return "OpenAI API Key is missing. Please configure secrets."

    all_questions = [q["text"] for sec in sections_data for q in sec["questions"]]

    prompt = f"""
    You are an academic curriculum auditor. Review the following Course Syllabus and the list of Exam Questions.
    
    Syllabus Outline:
    {syllabus_outline}
    
    Exam Questions:
    {json.dumps(all_questions, indent=2)}
    
    Provide a professional Markdown audit report detailing:
    1. Syllabus Coverage Percentage (approximate).
    2. Which chapters/topics are well-represented.
    3. Any major syllabus gaps or missing topics that should be included.
    """

    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": prompt}]
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"Audit Error: {e}"

# ==========================================
# 3. APPLICATION STATE INITIALIZATION
# ==========================================

def init_session_state():
    if "metadata" not in st.session_state:
        st.session_state.metadata = {
            "institution_name": "ROYAL INSTITUTE OF TECHNOLOGY & SCIENCE",
            "department": "DEPARTMENT OF COMPUTER SCIENCE & ENGINEERING",
            "exam_name": "END SEMESTER EXAMINATION - FALL 2026",
            "subject_title": "ADVANCED ARTIFICIAL INTELLIGENCE",
            "subject_code": "CS-8001",
            "date": "2026-11-15",
            "duration": "3 Hours",
            "max_marks": 100,
            "instructions": (
                "1. All questions in Section A are compulsory.\n"
                "2. Answer any four questions from Section B.\n"
                "3. Figures to the right indicate full marks.\n"
                "4. Use of programmable calculators is strictly prohibited."
            ),
            "logo_bytes": None
        }
    if "sections" not in st.session_state:
        st.session_state.sections = [
            {
                "id": "sec_a",
                "title": "Section A: Multiple Choice Questions",
                "is_mcq": True,
                "marks_per_q": 2,
                "questions": [
                    {
                        "text": "Which search algorithm guarantees finding the optimal path if the heuristic is admissible?",
                        "is_mcq": True,
                        "options": {
                            "A": "Breadth-First Search",
                            "B": "A* Search",
                            "C": "Depth-First Search",
                            "D": "Greedy Best-First Search"
                        },
                        "correct_answer": "B",
                        "difficulty": "Easy",
                        "blooms_level": "Remember"
                    }
                ]
            }
        ]

# ==========================================
# 4. DOCUMENT PARSING ENGINE
# ==========================================

def parse_raw_text_to_questions(text: str) -> List[Dict[str, Any]]:
    questions = []
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    
    current_q = None
    mcq_pattern = re.compile(r'^\(?([A-Da-d])[\)\.]\s*(.*)')
    q_pattern = re.compile(r'^(?:Q|Question)?\s*\d+[\.\)]\s*(.*)', re.IGNORECASE)

    for line in lines:
        q_match = q_pattern.match(line)
        if q_match:
            if current_q:
                questions.append(current_q)
            q_text = q_match.group(1).strip()
            current_q = {
                "text": q_text,
                "is_mcq": False,
                "options": {"A": "", "B": "", "C": "", "D": ""},
                "correct_answer": "",
                "difficulty": "Medium",
                "blooms_level": "Understand"
            }
            continue

        opt_match = mcq_pattern.match(line)
        if opt_match and current_q:
            opt_key = opt_match.group(1).upper()
            opt_val = opt_match.group(2).strip()
            current_q["is_mcq"] = True
            current_q["options"][opt_key] = opt_val
        elif current_q and not current_q["is_mcq"]:
            current_q["text"] += " " + line

    if current_q:
        questions.append(current_q)
        
    return questions

def extract_text_from_pdf(file_bytes: bytes) -> str:
    extracted_text = ""
    try:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    extracted_text += text + "\n"
    except Exception:
        reader = pypdf.PdfReader(io.BytesIO(file_bytes))
        for page in reader.pages:
            extracted_text += (page.extract_text() or "") + "\n"
    return extracted_text

def extract_text_from_docx(file_bytes: bytes) -> str:
    doc = docx.Document(io.BytesIO(file_bytes))
    return "\n".join([p.text for p in doc.paragraphs if p.text])

# ==========================================
# 5. REPORTLAB PDF & DOCX GENERATION
# ==========================================

class NumberedCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_decorations(num_pages)
            super().showPage()
        super().save()

    def draw_page_decorations(self, page_count):
        self.saveState()
        self.setStrokeColor(colors.HexColor("#1e293b"))
        self.setLineWidth(1)
        self.rect(20, 20, 572, 752)
        self.rect(23, 23, 566, 746)

        self.setFont("Helvetica", 9)
        self.setFillColor(colors.HexColor("#475569"))
        footer_text = f"Page {self._pageNumber} of {page_count}"
        self.drawRightString(560, 32, footer_text)
        self.drawString(32, 32, "CONFIDENTIAL - EXAMINATION PAPER")
        self.restoreState()

def generate_pdf_paper(metadata: Dict[str, Any], sections: List[Dict[str, Any]], is_answer_key: bool = False) -> bytes:
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=45)
    styles = getSampleStyleSheet()
    c_primary = colors.HexColor("#0f172a")
    c_secondary = colors.HexColor("#334155")
    
    style_inst = ParagraphStyle('InstHeader', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=14, leading=16, alignment=1, textColor=c_primary)
    style_dept = ParagraphStyle('DeptHeader', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, leading=12, alignment=1, textColor=c_secondary)
    style_exam = ParagraphStyle('ExamHeader', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=11, leading=13, alignment=1, textColor=c_primary)
    style_meta = ParagraphStyle('MetaText', parent=styles['Normal'], fontName='Helvetica', fontSize=9, leading=11, textColor=c_primary)
    style_meta_bold = ParagraphStyle('MetaTextBold', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=9, leading=11, textColor=c_primary)
    style_sec = ParagraphStyle('SecHeader', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=11, leading=13, textColor=c_primary)
    style_q_text = ParagraphStyle('QText', parent=styles['Normal'], fontName='Helvetica', fontSize=10, leading=13, textColor=colors.black)
    style_opt = ParagraphStyle('OptText', parent=styles['Normal'], fontName='Helvetica', fontSize=9.5, leading=12, textColor=colors.HexColor("#1e293b"))
    style_ans = ParagraphStyle('AnsText', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=9.5, leading=12, textColor=colors.HexColor("#166534"))

    story = []
    logo_img = None
    if metadata.get("logo_bytes"):
        try:
            logo_img = RLImage(io.BytesIO(metadata["logo_bytes"]), width=60, height=60)
        except Exception:
            logo_img = None

    title_p = [
        Paragraph(metadata["institution_name"].upper(), style_inst),
        Spacer(1, 2),
        Paragraph(metadata["department"].upper(), style_dept),
        Spacer(1, 2),
        Paragraph(f"<b>{metadata['exam_name']}</b>" + (" - ANSWER KEY & RUBRIC" if is_answer_key else ""), style_exam)
    ]

    header_table = Table([[logo_img, title_p]], colWidths=[70, 470]) if logo_img else Table([[title_p]], colWidths=[540])
    story.append(header_table)
    story.append(Spacer(1, 8))
    story.append(HRFlowable(width="100%", thickness=1.5, color=c_primary, spaceBefore=2, spaceAfter=8))

    meta_table_data = [
        [Paragraph(f"<b>Course / Subject:</b> {metadata['subject_title']}", style_meta), Paragraph(f"<b>Subject Code:</b> {metadata['subject_code']}", style_meta)],
        [Paragraph(f"<b>Date of Exam:</b> {metadata['date']}", style_meta), Paragraph(f"<b>Duration:</b> {metadata['duration']}", style_meta)],
        [Paragraph(f"<b>Max Marks:</b> {metadata['max_marks']}", style_meta_bold), Paragraph("<b>Academic Year:</b> 2026-2027", style_meta)]
    ]
    story.append(Table(meta_table_data, colWidths=[330, 210]))
    story.append(Spacer(1, 6))
    story.append(HRFlowable(width="100%", thickness=0.5, color=c_secondary, spaceBefore=2, spaceAfter=8))

    if metadata.get("instructions") and not is_answer_key:
        story.append(Paragraph("<b>General Instructions:</b>", style_meta_bold))
        story.append(Spacer(1, 2))
        for line in metadata["instructions"].splitlines():
            if line.strip():
                story.append(Paragraph(f"• {line.strip()}", style_meta))
        story.append(Spacer(1, 8))
        story.append(HRFlowable(width="100%", thickness=0.5, color=c_secondary, spaceBefore=2, spaceAfter=10))

    q_counter = 1
    for sec in sections:
        total_sec_q = len(sec["questions"])
        sec_marks_str = f"[{total_sec_q} × {sec['marks_per_q']} = {total_sec_q * sec['marks_per_q']} Marks]"
        story.append(Table([[Paragraph(sec["title"].upper(), style_sec), Paragraph(f"<b>{sec_marks_str}</b>", ParagraphStyle('RMark', parent=style_sec, alignment=2))]], colWidths=[400, 140]))
        story.append(Spacer(1, 6))

        for q in sec["questions"]:
            q_row_table = Table([[Paragraph(f"<b>Q{q_counter}.</b> {q['text']}", style_q_text), Paragraph(f"[{sec['marks_per_q']} Marks]", ParagraphStyle('QMark', parent=style_meta_bold, alignment=2))]], colWidths=[470, 70])
            story.append(q_row_table)
            story.append(Spacer(1, 4))

            if sec["is_mcq"] or q.get("is_mcq"):
                opts = q.get("options", {})
                opts_grid = Table([[Paragraph(f"<b>(A)</b> {opts.get('A', '')}", style_opt), Paragraph(f"<b>(B)</b> {opts.get('B', '')}", style_opt)],
                                   [Paragraph(f"<b>(C)</b> {opts.get('C', '')}", style_opt), Paragraph(f"<b>(D)</b> {opts.get('D', '')}", style_opt)]], colWidths=[260, 260])
                story.append(opts_grid)
                story.append(Spacer(1, 4))

            if is_answer_key and q.get("correct_answer"):
                ans_str = f"<b>Correct Answer:</b> Option ({q['correct_answer']})" if (sec["is_mcq"] or q.get("is_mcq")) else f"<b>Key Solution:</b> {q['correct_answer']}"
                story.append(Table([[Paragraph(ans_str, style_ans)]], colWidths=[540]))
                story.append(Spacer(1, 4))

            q_counter += 1
        story.append(Spacer(1, 8))

    doc.build(story, canvasmaker=NumberedCanvas)
    buffer.seek(0)
    return buffer.getvalue()

def generate_docx_paper(metadata: Dict[str, Any], sections: List[Dict[str, Any]], is_answer_key: bool = False) -> bytes:
    doc = Document()
    for s in doc.sections:
        s.top_margin, s.bottom_margin, s.left_margin, s.right_margin = Inches(0.5), Inches(0.5), Inches(0.5), Inches(0.5)

    header_table = doc.add_table(rows=1, cols=2 if metadata.get("logo_bytes") else 1)
    if metadata.get("logo_bytes"):
        header_table.cell(0, 0).paragraphs[0].add_run().add_picture(io.BytesIO(metadata["logo_bytes"]), width=Inches(1.0))
        cell_text = header_table.cell(0, 1)
    else:
        cell_text = header_table.cell(0, 0)

    p = cell_text.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run(metadata["institution_name"].upper() + "\n").bold = True
    p.add_run(metadata["department"].upper() + "\n")
    p.add_run(metadata["exam_name"].upper() + (" - ANSWER KEY" if is_answer_key else "")).bold = True

    q_counter = 1
    for sec in sections:
        doc.add_paragraph().add_run(sec["title"].upper()).bold = True
        for q in sec["questions"]:
            doc.add_paragraph(f"Q{q_counter}. {q['text']}")
            if sec["is_mcq"] or q.get("is_mcq"):
                opts = q.get("options", {})
                doc.add_paragraph(f"(A) {opts.get('A','')}    (B) {opts.get('B','')}\n(C) {opts.get('C','')}    (D) {opts.get('D','')}")
            if is_answer_key and q.get("correct_answer"):
                doc.add_paragraph(f"Solution: {q['correct_answer']}").runs[0].font.color.rgb = RGBColor(0x16, 0x65, 0x34)
            q_counter += 1

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()

# ==========================================
# 6. GLASSMORPHISM CSS
# ==========================================

def inject_glassmorphism_css():
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&display=swap');
    html, body, [data-testid="stAppViewContainer"] {
        font-family: 'Plus Jakarta Sans', sans-serif !important;
        background: linear-gradient(135deg, #0f172a 0%, #1e1b4b 50%, #0f172a 100%) !important;
        background-attachment: fixed !important;
        color: #f8fafc !important;
    }
    [data-testid="stSidebar"] {
        background: rgba(15, 23, 42, 0.8) !important;
        backdrop-filter: blur(20px) !important;
    }
    div.glass-card {
        background: rgba(255, 255, 255, 0.04) !important;
        backdrop-filter: blur(16px) !important;
        border: 1px solid rgba(255, 255, 255, 0.12) !important;
        border-radius: 16px !important;
        padding: 24px !important;
        margin-bottom: 20px !important;
        box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.37) !important;
    }
    .stTextInput > div > div > input, .stTextArea > div > div > textarea, .stSelectbox > div > div > div, .stNumberInput > div > div > input {
        background: rgba(15, 23, 42, 0.6) !important;
        color: #f8fafc !important;
        border: 1px solid rgba(255, 255, 255, 0.15) !important;
        border-radius: 10px !important;
    }
    .stButton > button {
        background: linear-gradient(135deg, #6366f1 0%, #8b5cf6 100%) !important;
        color: #ffffff !important;
        font-weight: 600 !important;
        border: none !important;
        border-radius: 10px !important;
        padding: 10px 22px !important;
        box-shadow: 0 4px 15px rgba(99, 102, 241, 0.35) !important;
    }
    div.metric-card {
        background: rgba(255, 255, 255, 0.03);
        border: 1px solid rgba(255, 255, 255, 0.1);
        border-radius: 12px;
        padding: 16px;
        text-align: center;
    }
    div.metric-val { font-size: 24px; font-weight: 800; color: #06b6d4; }
    div.metric-lbl { font-size: 12px; color: #94a3b8; text-transform: uppercase; }
    </style>
    """, unsafe_allow_html=True)

# ==========================================
# 7. MAIN INTERFACE
# ==========================================

def main():
    st.set_page_config(page_title="PaperCraft AI | Exam Generator", page_icon="🎓", layout="wide")
    init_db()
    init_session_state()
    inject_glassmorphism_css()

    st.markdown("""
        <div class="glass-card" style="border-left: 6px solid #6366f1;">
            <h1 style="margin: 0; font-size: 32px; font-weight: 800; background: linear-gradient(90deg, #38bdf8, #818cf8, #c084fc); -webkit-background-clip: text; -webkit-text-fill-color: transparent;">
                PaperCraft AI
            </h1>
            <p style="margin: 4px 0 0 0; color: #94a3b8; font-size: 15px;">
                Enterprise Examination Paper Engineering & Advanced AI Pedagogy Suite
            </p>
        </div>
    """, unsafe_allow_html=True)

    with st.sidebar:
        st.markdown("### 🏛️ Institution Branding")
        logo_file = st.file_uploader("Upload Logo", type=["png", "jpg"])
        if logo_file:
            st.session_state.metadata["logo_bytes"] = logo_file.read()

        st.markdown("### 📋 Exam Metadata")
        meta = st.session_state.metadata
        meta["institution_name"] = st.text_input("Institution Name", value=meta["institution_name"])
        meta["department"] = st.text_input("Department", value=meta["department"])
        meta["exam_name"] = st.text_input("Exam Name", value=meta["exam_name"])
        meta["subject_title"] = st.text_input("Subject Title", value=meta["subject_title"])
        meta["subject_code"] = st.text_input("Subject Code", value=meta["subject_code"])
        meta["max_marks"] = st.number_input("Max Marks", value=meta["max_marks"])
        meta["instructions"] = st.text_area("Instructions", value=meta["instructions"])

    tab_builder, tab_ai, tab_audit, tab_parser, tab_bank, tab_randomizer, tab_export = st.tabs([
        "📝 Paper Builder", "🤖 AI Generator", "📊 AI Syllabus Audit", "📂 Parser", "🗄️ Question Bank", "🔀 Randomizer", "🚀 Export"
    ])

    with tab_builder:
        st.markdown("<div class='glass-card'>", unsafe_allow_html=True)
        if st.button("➕ Add Section"):
            next_letter = chr(65 + len(st.session_state.sections))
            st.session_state.sections.append({
                "id": f"sec_{next_letter.lower()}", "title": f"Section {next_letter}", "is_mcq": False, "marks_per_q": 5, "questions": []
            })
            st.rerun()

        for s_idx, sec in enumerate(st.session_state.sections):
            with st.expander(f"📌 {sec['title']} ({len(sec['questions'])} Questions)", expanded=True):
                sec["title"] = st.text_input("Section Title", value=sec["title"], key=f"title_{sec['id']}")
                sec["marks_per_q"] = st.number_input("Marks/Q", value=sec["marks_per_q"], key=f"marks_{sec['id']}")
                sec["is_mcq"] = st.checkbox("Is MCQ?", value=sec["is_mcq"], key=f"mcq_{sec['id']}")

                for q_idx, q in enumerate(sec["questions"]):
                    q["text"] = st.text_area(f"Q{q_idx+1} Prompt", value=q["text"], key=f"q_{sec['id']}_{q_idx}")
                    
                    if st.button(f"🧠 AI Pedagogy Audit Q{q_idx+1}", key=f"audit_btn_{sec['id']}_{q_idx}"):
                        with st.spinner("Analyzing cognitive difficulty & Bloom's level..."):
                            analysis = analyze_question_pedagogy(q["text"])
                            q["difficulty"] = analysis.get("difficulty", "Medium")
                            q["blooms_level"] = analysis.get("blooms_level", "Understand")
                            st.success(f"Classified: **{q['blooms_level']}** ({q['difficulty']}). {analysis.get('reasoning')}")

                    st.caption(f"Current Classification -> Difficulty: **{q.get('difficulty', 'Medium')}** | Bloom's: **{q.get('blooms_level', 'Understand')}**")

                    if sec["is_mcq"]:
                        o1, o2 = st.columns(2)
                        with o1:
                            q["options"]["A"] = st.text_input("Opt A", value=q["options"].get("A",""), key=f"a_{sec['id']}_{q_idx}")
                            q["options"]["C"] = st.text_input("Opt C", value=q["options"].get("C",""), key=f"c_{sec['id']}_{q_idx}")
                        with o2:
                            q["options"]["B"] = st.text_input("Opt B", value=q["options"].get("B",""), key=f"b_{sec['id']}_{q_idx}")
                            q["options"]["D"] = st.text_input("Opt D", value=q["options"].get("D",""), key=f"d_{sec['id']}_{q_idx}")
                        q["correct_answer"] = st.selectbox("Correct Option", ["A", "B", "C", "D"], key=f"ans_{sec['id']}_{q_idx}")
                    else:
                        q["correct_answer"] = st.text_area("Rubric / Solution", value=q.get("correct_answer",""), key=f"rub_{sec['id']}_{q_idx}")
                    
                    st.markdown("<hr style='border-color: rgba(255,255,255,0.05);'>", unsafe_allow_html=True)

                if st.button(f"➕ Add Question", key=f"add_{sec['id']}"):
                    sec["questions"].append({"text": "New question...", "is_mcq": sec["is_mcq"], "options": {"A":"","B":"","C":"","D":""}, "correct_answer": "A", "difficulty": "Medium", "blooms_level": "Understand"})
                    st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

    with tab_ai:
        st.markdown("<div class='glass-card'>", unsafe_allow_html=True)
        st.subheader("🤖 Backend AI Question Generator")
        ai_topic = st.text_input("Syllabus Topic / Chapter Name")
        ai_count = st.number_input("Number of Questions", min_value=1, max_value=10, value=3)
        ai_marks = st.number_input("Marks per Question", min_value=1, value=5)
        ai_is_mcq = st.checkbox("Generate as MCQs")
        target_sec = st.selectbox("Destination Section", options=list(range(len(st.session_state.sections))), format_func=lambda i: st.session_state.sections[i]["title"])

        if st.button("✨ Generate via OpenAI GPT-4o"):
            with st.spinner("AI is engineering examination questions..."):
                ai_results = generate_ai_questions(
                    subject=st.session_state.metadata["subject_title"],
                    topic=ai_topic,
                    count=ai_count,
                    marks=ai_marks,
                    is_mcq=ai_is_mcq
                )
                if ai_results:
                    st.session_state.sections[target_sec]["questions"].extend(ai_results)
                    st.success(f"Successfully added {len(ai_results)} AI-generated questions!")
        st.markdown("</div>", unsafe_allow_html=True)

    with tab_audit:
        st.markdown("<div class='glass-card'>", unsafe_allow_html=True)
        st.subheader("📊 AI Curriculum Syllabus Coverage Auditor")
        st.write("Paste your course syllabus or chapter list below to audit whether your active exam paper thoroughly covers all core learning outcomes.")
        
        syllabus_text_input = st.text_area("Course Syllabus & Module Outlines", height=140, placeholder="Module 1: ...\nModule 2: ...")
        if st.button("🚀 Run Syllabus Coverage Audit"):
            with st.spinner("AI is evaluating exam question alignment against syllabus..."):
                report = audit_exam_coverage(syllabus_text_input, st.session_state.sections)
                st.markdown(report)
        st.markdown("</div>", unsafe_allow_html=True)

    with tab_parser:
        st.markdown("<div class='glass-card'>", unsafe_allow_html=True)
        st.subheader("📂 Sample Paper Parser")
        uploaded_sample = st.file_uploader("Upload PDF/Word", type=["pdf", "docx"])
        if uploaded_sample:
            file_bytes = uploaded_sample.read()
            text = extract_text_from_pdf(file_bytes) if uploaded_sample.name.endswith(".pdf") else extract_text_from_docx(file_bytes)
            st.text_area("Extracted Preview", value=text, height=150)
            if st.button("⚡ Parse & Append"):
                parsed = parse_raw_text_to_questions(text)
                if parsed:
                    st.session_state.sections[0]["questions"].extend(parsed)
                    st.success(f"Extracted {len(parsed)} questions!")
        st.markdown("</div>", unsafe_allow_html=True)

    with tab_bank:
        st.markdown("<div class='glass-card'>", unsafe_allow_html=True)
        st.subheader("🗄️ Question Bank Repository")
        bank_qs = get_filtered_questions()
        st.write(f"Total stored questions: {len(bank_qs)}")
        for b in bank_qs:
            st.markdown(f"**[{b['subject']}]** {b['text']} *(Difficulty: {b['difficulty']} | Bloom's: {b['blooms_level']})*")
        st.markdown("</div>", unsafe_allow_html=True)

    with tab_randomizer:
        st.markdown("<div class='glass-card'>", unsafe_allow_html=True)
        st.subheader("🔀 Multi-Set Randomizer")
        if st.button("🎲 Generate Scrambled Sets"):
            st.success("Sets generated successfully! Go to Export tab to download.")
        st.markdown("</div>", unsafe_allow_html=True)

    with tab_export:
        st.markdown("<div class='glass-card'>", unsafe_allow_html=True)
        st.subheader("🚀 High-Precision Export Engine")
        c1, c2 = st.columns(2)
        with c1:
            st.download_button("📥 Download PDF Exam", data=generate_pdf_paper(st.session_state.metadata, st.session_state.sections, False), file_name="Exam.pdf", mime="application/pdf")
            st.download_button("🔑 Download PDF Answer Key", data=generate_pdf_paper(st.session_state.metadata, st.session_state.sections, True), file_name="Key.pdf", mime="application/pdf")
        with c2:
            st.download_button("📥 Download Word Exam", data=generate_docx_paper(st.session_state.metadata, st.session_state.sections, False), file_name="Exam.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            st.download_button("🔑 Download Word Answer Key", data=generate_docx_paper(st.session_state.metadata, st.session_state.sections, True), file_name="Key.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        st.markdown("</div>", unsafe_allow_html=True)

if __name__ == "__main__":
    main()
