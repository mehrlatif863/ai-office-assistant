import streamlit as st
import re
import io
import csv
import pandas as pd
import plotly.express as px # <-- NEW: For fancy charts
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
# 3. CREATE TABS (Added Tab 3)
# ==========================================
tab1, tab2, tab3 = st.tabs(["💬 AI Chat", "📊 Auto Insights", "🏢 Executive Dashboard"])

# ==========================================
# TAB 1: AI CHATBOT (HIDDEN TO SAVE SPACE, BUT STILL WORKS)
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
        words = set(re.findall(r'\b\w+\b', q.lower())) - {"what","is","the","a","an","in","on","to","for","of","with","how","does","do","did","are","was","were","be","this","that","it","from","by","and","or","but","if"}
        if not words: return "\n".join(chunks[:3])
        scores = sorted([(sum(1 for w in words if w in c.lower()), c) for c in chunks], key=lambda x: x[0], reverse=True)
        return "\n".join([c for s, c in scores[:3]])

    if uploaded_files:
        all_text = ""
        with st.spinner(f"Reading {len(uploaded_files)} files..."):
            for file in uploaded_files:
                ft = file.name.split(".")[-1].lower()
                if ft=="pdf": all_text += read_pdf(file) + "\n\n"
                elif ft=="docx": all_text += read_word(file) + "\n\n"
                elif ft=="pptx": all_text += read_pptx(file) + "\n\n"
                elif ft=="csv": all_text += read_csv_txt(file) + "\n\n"
                elif ft=="xlsx": all_text += read_excel_txt(file) + "\n\n"
            chunks = chunk_text(all_text)
        st.success(f"✅ {len(uploaded_files)} files indexed!")
        for msg in st.session_state.messages:
            with st.chat_message(msg["role"]): st.markdown(msg["content"])
        q = st.chat_input("Ask across all files...")
        if q:
            st.session_state.messages.append({"role": "user", "content": q})
            with st.chat_message("user"): st.markdown(q)
            with st.chat_message("assistant"):
                with st.spinner("Thinking..."):
                    ctx = find_best_chunks(chunks, q)
                    res = client.chat.completions.create(model="llama-3.3-70b-versatile", messages=[{"role":"user","content":f"Answer ONLY from context:\n{ctx}\n\nQ:{q}"}])
                    ans = res.choices[0].message.content
                    st.markdown(ans)
            st.session_state.messages.append({"role": "assistant", "content": ans})
            st.session_state.last_answer = ans
        if st.session_state.last_answer:
            if st.button("📥 Download Answer as Word"):
                doc = DocxDocument(); doc.add_paragraph(st.session_state.last_answer)
                buf = io.BytesIO(); doc.save(buf); buf.seek(0)
                st.download_button("Save File", buf, "Answer.docx")

# ==========================================
# TAB 2: AUTO DASHBOARD (HIDDEN BUT WORKS)
# ==========================================
with tab2:
    dash_file = st.file_uploader("Upload Data File", type=["xlsx", "csv"], key="dash")
    if dash_file:
        df = pd.read_csv(dash_file) if dash_file.name.endswith(".csv") else pd.read_excel(dash_file)
        st.success("✅ Data loaded!")
        cols = df.columns.tolist()
        if len(cols) >= 2:
            c1, c2 = st.columns(2)
            with c1: x_axis = st.selectbox("X-Axis", cols, key="x_ax")
            with c2: y_axis = st.selectbox("Y-Axis", cols, index=1, key="y_ax")
            if st.button("📊 Generate Chart", use_container_width=True):
                try: st.bar_chart(df, x=x_axis, y=y_axis)
                except: st.error("Ensure Y-Axis has numbers.")

# ==========================================
# TAB 3: EXECUTIVE DASHBOARD (THE NEW MAGIC)
# ==========================================
with tab3:
    st.markdown("### 📈 Store Performance Overview")
    exec_file = st.file_uploader("Upload Store Data (Excel/CSV)", type=["xlsx", "csv"], key="exec")
    
    if exec_file:
        df = pd.read_csv(exec_file) if exec_file.name.endswith(".csv") else pd.read_excel(exec_file)
        
        # Clean column names (remove spaces to make it easier)
        df.columns = df.columns.str.replace(' ', '_')
        
        # Try to convert date column
        date_col = None
        for col in df.columns:
            if 'date' in col.lower() or 'time' in col.lower():
                df[col] = pd.to_datetime(df[col], errors='coerce')
                date_col = col
                break
                
        # --- PART 1: KPI METRICS ---
        st.markdown("---")
        
        # Find Sales/Cost columns dynamically
        sales_col = next((c for c in df.columns if 'sales' in c.lower() or 'revenue' in c.lower()), None)
        cost_col = next((c for c in df.columns if 'cost' in c.lower() or 'purchase' in c.lower() or 'expense' in c.lower()), None)
        
        kpi1, kpi2, kpi3, kpi4 = st.columns(4)
        
        if sales_col:
            total_sales = df[sales_col].sum()
            kpi1.metric(label="💰 Total Sales", value=f"${total_sales:,.0f}")
            
            if cost_col:
                total_cost = df[cost_col].sum()
                profit = total_sales - total_cost
                kpi2.metric(label="📉 Total Cost", value=f"${total_cost:,.0f}")
                kpi3.metric(label="✨ Net Profit", value=f"${profit:,.0f}")
        
        # Find Top Customer and City
        cust_col = next((c for c in df.columns if 'customer' in c.lower() or 'client' in c.lower()), None)
        city_col = next((c for c in df.columns if 'city' in c.lower() or 'location' in c.lower() or 'state' in c.lower()), None)
        
        if cust_col and sales_col:
            top_cust = df.groupby(cust_col)[sales_col].sum().idxmax()
            kpi4.metric(label="🏆 Top Customer", value=top_cust)
        elif city_col and sales_col:
            top_city = df.groupby(city_col)[sales_col].sum().idxmax()
            kpi4.metric(label="📍 Top Location", value=top_city)

        # --- PART 2: CHARTS GRID ---
        st.markdown("---")
        chart1, chart2 = st.columns(2)
        
        # Chart 1: Sales Trend (Line)
        with chart1:
            st.markdown("**Sales Trend Over Time**")
            if date_col and sales_col:
                trend_data = df.groupby(date_col)[sales_col].sum().reset_index()
                fig = px.line(trend_data, x=date_col, y=sales_col, template="plotly_white")
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.warning("Need a 'Date' and 'Sales' column for trend.")

        # Chart 2: Top Customers (Bar)
        with chart2:
            st.markdown("**Top Customers**")
            if cust_col and sales_col:
                cust_data = df.groupby(cust_col)[sales_col].sum().sort_values(ascending=False).head(5).reset_index()
                fig = px.bar(cust_data, x=cust_col, y=sales_col, template="plotly_white")
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.warning("Need a 'Customer' and 'Sales' column.")

        chart3, chart4 = st.columns(2)
        
        # Chart 3: Sales by Location (Donut)
        with chart3:
            st.markdown("**Sales by Location**")
            if city_col and sales_col:
                city_data = df.groupby(city_col)[sales_col].sum().reset_index()
                fig = px.pie(city_data, values=sales_col, names=city_col, hole=0.4, template="plotly_white")
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.warning("Need a 'City/Location' column.")

        # Chart 4: Sales by Category (Donut)
        with chart4:
            st.markdown("**Sales by Category**")
            cat_col = next((c for c in df.columns if 'category' in c.lower() or 'product' in c.lower() or 'item' in c.lower()), None)
            if cat_col and sales_col:
                cat_data = df.groupby(cat_col)[sales_col].sum().reset_index()
                fig = px.pie(cat_data, values=sales_col, names=cat_col, hole=0.4, template="plotly_white")
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.warning("Need a 'Category/Product' column.")

# ==========================================
# SIDEBAR
# ==========================================
with st.sidebar:
    if st.button("🗑️ Clear Chat History"): st.session_state.messages = []; st.rerun()
    if st.button("🔒 Logout"): st.session_state.authenticated = False; st.rerun()
