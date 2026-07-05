import streamlit as st
import re
import io
import csv
from groq import Groq
from pypdf import PdfReader
from docx import Document as DocxDocument # Changed to avoid conflict
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

# Keep track of chat history and current answer for export
if "messages" not in st.session_state:
    st.session_state.messages = []
if "last_answer" not in st.session_state:
    st.session_state.last_answer = ""
if "last_question" not in st.session_state:
    st.session_state.last_question = ""

st.title("🤖 AI Office Assistant (Pro)")
st.write("Chat with your documents. Export answers to Word.")

# ==========================================
# 3. FILE READERS (Updated Word Reader)
# ==========================================
def read_pdf(file):
    text = ""
    for page in PdfReader(file).pages: text += page.extract_text() + "\n"
    return text

def read_word(file):
    text = ""
    doc = DocxDocument(file) # Updated name
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
    def read_csv(file):
        text =""
    # Handle different text encodings
    try:
        decoded_file = io.StringIO(file.read().decode('utf-8'))
    except UnicodeDecodeError:
        decoded_file = io.StringIO(file.read().decode('latin-1'))
        reader = csv.reader(decoded_file)
    for row in reader:
        text += " | ".join(row) + "\n"
    return text
# ==========================================
# 4. RAG LOGIC
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
# 5. LAYOUT: CHAT HISTORY SIDEBAR
# ==========================================
with st.sidebar:
    st.markdown("### 💬 Chat History")
    if len(st.session_state.messages) == 0:
        st.info("No messages yet.")
    else:
        for msg in st.session_state.messages:
            if msg["role"] == "user":
                st.markdown(f"**👤 You:** {msg['content'][:50]}...")
            else:
                st.markdown(f"**🤖 AI:** {msg['content'][:50]}...")
    
    st.markdown("---")
    if st.button("🗑️ Clear Chat History"):
        st.session_state.messages = []
        st.rerun()
        
    if st.button("🔒 Logout"):
        st.session_state.authenticated = False
        st.rerun()

# ==========================================
# 6. MAIN APP LOGIC
# ==========================================
uploaded_file = st.file_uploader("Upload Document", type=["pdf", "docx", "xlsx", "pptx","csv"])

if uploaded_file is not None:
    file_type = uploaded_file.name.split(".")[-1].lower()
    
    with st.spinner("Reading & Indexing..."):
        if file_type == "pdf": 
            doc_text = read_pdf(uploaded_file)
        elif file_type == "docx": 
            doc_text = read_word(uploaded_file)
        elif file_type == "xlsx": 
            doc_text = read_excel(uploaded_file)
        elif file_type == "pptx": 
            doc_text = read_pptx(uploaded_file)
        elif file_type == "csv": 
            doc_text = read_csv(uploaded_file) 
        else: 
            doc_text = ""
            
        document_chunks = chunk_text(doc_text)
        st.success(f"✅ {uploaded_file.name} ready! ({len(document_chunks)} segments)")

    # Display past messages
    for msg in st.session_state.messages:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    question = st.chat_input("Ask about your file...")
    
    if question:
        # 1. Save to history & display user question
        st.session_state.messages.append({"role": "user", "content": question})
        with st.chat_message("user"):
            st.markdown(question)
            
        # 2. Get AI answer
        with st.chat_message("assistant"):
            with st.spinner("Searching & thinking..."):
                best_context = find_best_chunks(document_chunks, question)
                prompt = f"Answer based ONLY on this context:\n{best_context}\n\nQuestion: {question}"
                response = client.chat.completions.create(
                    model="llama-3.3-70b-versatile", 
                    messages=[{"role": "user", "content": prompt}]
                )
                answer = response.choices[0].message.content
                st.markdown(answer)
                
        # 3. Save to history & session state for export
        st.session_state.messages.append({"role": "assistant", "content": answer})
        st.session_state.last_question = question
        st.session_state.last_answer = answer

    # ==========================================
    # 7. EXPORT TO WORD FEATURE
    # ==========================================
    if st.session_state.last_answer:
        st.markdown("---")
        col1, col2, col3 = st.columns([1,1,1])
        with col2:
            if st.button("📥 Download Last Answer as Word Document", use_container_width=True):
                # Create Word Document in memory
                doc = DocxDocument()
                doc.add_heading('AI Office Assistant - Export', 0)
                doc.add_heading('Your Question:', level=1)
                doc.add_paragraph(st.session_state.last_question)
                doc.add_heading('AI Answer:', level=1)
                doc.add_paragraph(st.session_state.last_answer)
                
                # Save to memory buffer
                buffer = io.BytesIO()
                doc.save(buffer)
                buffer.seek(0)
                
                # Trigger download
                st.download_button(
                    label="Click here to save the .docx file",
                    data=buffer,
                    file_name="AI_Assistant_Answer.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
