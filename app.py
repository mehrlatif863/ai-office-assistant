import streamlit as st
import re
import io
import csv
import pandas as pd
import plotly.express as px
from groq import Groq
from pypdf import PdfReader
from docx import Document as DocxDocument
import openpyxl
from pptx import Presentation

# ==========================================
# 1. SECURITY CHECK
# ==========================================
def get_real_key():
    p1 = "M."
    p2 = "LATIF"
    p3 = "PPC"
    return p1 + p2 + p3

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.markdown("<h1 style='text-align: center;'>🔒 Secure AI Office Assistant</h1>", unsafe_allow_html=True)
    with st.form("login"):
        user_key = st.text_input("Enter Access Key:", type="password")
        if st.form_submit_button("Unlock"):
            if user_key == get_real_key():
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Invalid Key!")
    st.stop()

# ==========================================
# 2. SETUP AI & SESSION STATE
# ==========================================
client = Groq(api_key=st.secrets["GROQ_API_KEY"])
if "messages" not in st.session_state:
    st.session_state.messages = []
if "last_answer" not in st.session_state:
    st.session_state.last_answer = ""

st.title("🤖 AI Office Assistant (Enterprise Edition)")

# ==========================================
# 3. CREATE TABS
# ==========================================
tab1, tab2, tab3 = st.tabs(["💬 AI Chat", "📊 Auto Insights", "🏢 Executive Dashboard"])

# ==========================================
# TAB 1: AI CHATBOT
# ==========================================
with tab1:
    uploaded_files = st.file_uploader("Upload Documents", type=["pdf", "docx", "xlsx", "pptx", "csv"], accept_multiple_files=True, key="chat")
    
    def read_pdf(file): return "\n".join([p.extract_text() for p in PdfReader(file).pages])
    def read_word(file):
        doc = DocxDocument(file); text = "\n".join([p.text for p in doc.paragraphs])
        for t in doc.tables:
            for r in t.rows: text += "\n" + " | ".join([c.text for c in r.cells])
        return text
    def read_pptx(file):
        text = ""
        for s in Presentation(file).slides:
            for sh in s.shapes:
                if sh.has_text_frame: text += sh.text_frame.text + "\n"
        return text
    def read_csv_txt(file):
        try: f = io.StringIO(file.read().decode('utf-8'))
        except: f = io.StringIO(file.read().decode('latin-1'))
        return "\n".join([",".join(row) for row in csv.reader(f)])
    def read_excel_txt(file):
        wb = openpyxl.load_workbook(file); text = ""
        for sn in wb.sheetnames:
            text += f"--- {sn} ---\n"
            for row in wb[sn].iter_rows(values_only=True): text += " | ".join([str(c) if c else "" for c in row]) + "\n"
        return text

    def chunk_text(text, size=4000, overlap=400):
        chunks, start = [], 0
        while start < len(text): chunks.append(text[start:start+size]); start += size - overlap
        return chunks
    def find_best_chunks(chunks, q):
        words =
