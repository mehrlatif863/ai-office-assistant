import streamlit as st
import re
import io
import csv
import pandas as pd # <-- NEW: The data engine!
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
if "last_question" not in st.session_state:
    st.session_state.last_question = ""

st.title("🤖 AI Office Assistant (Data Analyst Edition)")
st.write("Chat with documents OR visualize your Excel/CSV data instantly.")

# ==========================================
# 3. FILE READERS
# ==========================================
def read_pdf(file):
    text = ""
    for page in PdfReader(file).pages: text += page.extract_text() + "\n"
    return text

def read_word(file):
    text = ""
    doc = DocxDocument(file)
    for para in doc.paragraphs: text += para.text + "\n"
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells: text += cell.text + " | "
            text += "\n"
    return text

def read_pptx(file):
    text = ""
    for slide_num, slide in enumerate(Presentation(file).slides, 1):
        text += f"--- Slide {slide_num} ---\n"
        for shape in slide.shapes:
            if shape.has_text_frame: text += shape.text_frame.text + "\n"
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells: text += cell.text + " | "
                    text += "\n"
    return text

# ==========================================
# 4. RAG LOGIC (For Chat Tab)
# ==========================================
def chunk_text(text, size=4000, overlap=400):
    chunks, start = [], 0
    while start < len(text):
        chunks.append(text[start:start + size])
        start += size - overlap
    return chunks

def find_best_chunks(chunks, question, top_k=3):
    words = set(re.findall(r'\b\w+\b', question.lower()))
    boring = {"what","is","the","a","an","in","on","to","for","of","with","how","does","do","did","are","was","were","be","this","that","it","from","by","and","or","but","if"}
    keywords = words - boring
    if not keywords: return "\n\n".join(chunks[:top_k])
    scores = [(sum(1 for w in keywords if w in c.lower()), c) for c in chunks]
    scores.sort(key=lambda x: x[0], reverse=True)
    return "\n\n".join([c for s, c in scores[:top_k]])

# ==========================================
# 5. LAYOUT: SIDEBAR
# ==========================================
with st.sidebar:
    st.markdown("### 💬 Chat History")
    if len(st.session_state.messages) == 0:
        st.info("No messages yet.")
    else:
        for msg in st.session_state.messages:
            if msg["role"] == "user":
                st.markdown(f"**👤
