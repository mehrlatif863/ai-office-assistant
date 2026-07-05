import streamlit as st
import re
from groq import Groq
from pypdf import PdfReader
from docx import Document
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
# 2. SETUP GROQ AI (Llama 3)
# ==========================================
# PASTE YOUR GROQ KEY HERE
client = Groq(api_key="client = Groq(api_key=st.secrets["GROQ_API_KEY"])")

st.title("🤖 AI Office Assistant (Groq Powered)")
st.write("Lightning fast. No limits.")

# ==========================================
# 3. FILE READERS
# ==========================================
def read_pdf(file):
    text = ""
    for page in PdfReader(file).pages: text += page.extract_text() + "\n"
    return text

def read_word(file):
    text = ""
    doc = Document(file)
    for para in doc.paragraphs: text += para.text + "\n"
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells: text += cell.text + " | "
            text += "\n"
    return text

def read_excel(file):
    text = ""
    wb = openpyxl.load_workbook(file)
    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        text += f"--- Sheet: {sheet_name} ---\n"
        for row in sheet.iter_rows(values_only=True):
            text += " | ".join([str(cell) if cell else "" for cell in row]) + "\n"
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
# 4. KEYWORD SEARCH RAG
# ==========================================
def chunk_text(text, size=4000, overlap=400):
    chunks = []
    start = 0
    while start < len(text):
        chunks.append(text[start:start + size])
        start += size - overlap
    return chunks

def find_best_chunks(chunks, question, top_k=3):
    words = set(re.findall(r'\b\w+\b', question.lower()))
    boring_words = {"what", "is", "the", "a", "an", "in", "on", "to", "for", "of", "with", "how", "does", "do", "did", "are", "was", "were", "be", "this", "that", "it", "from", "by", "and", "or", "but", "if"}
    keywords = words - boring_words
    if not keywords: return "\n\n".join(chunks[:top_k])
        
    scores = []
    for chunk in chunks:
        chunk_lower = chunk.lower()
        score = sum(1 for word in keywords if word in chunk_lower)
        scores.append((score, chunk))
        
    scores.sort(key=lambda x: x[0], reverse=True)
    return "\n\n".join([chunk for score, chunk in scores[:top_k]])

# ==========================================
# 5. MAIN APP LOGIC
# ==========================================
uploaded_file = st.file_uploader("Upload Document", type=["pdf", "docx", "xlsx", "pptx"])

if uploaded_file is not None:
    file_type = uploaded_file.name.split(".")[-1].lower()
    
    with st.spinner("Reading & Indexing..."):
        if file_type == "pdf": doc_text = read_pdf(uploaded_file)
        elif file_type == "docx": doc_text = read_word(uploaded_file)
        elif file_type == "xlsx": doc_text = read_excel(uploaded_file)
        elif file_type == "pptx": doc_text = read_pptx(uploaded_file)
        else: doc_text = ""
        
        document_chunks = chunk_text(doc_text)

    st.success(f"✅ {uploaded_file.name} ready! ({len(document_chunks)} segments created)")

    question = st.chat_input("Ask about your file...")
    
    if question:
        with st.spinner("Searching & thinking..."):
            best_context = find_best_chunks(document_chunks, question)
            
            prompt = f"Answer based ONLY on this context:\n{best_context}\n\nQuestion: {question}"
            
            # GROQ API CALL
            response = client.chat.completions.create(
                model="llama-3.3-70b-versatile", 
                messages=[{"role": "user", "content": prompt}]
            )
            st.write(response.choices[0].message.content)

    if st.button("🔒 Logout"):
        st.session_state.authenticated = False
        st.rerun()