import streamlit as st
import pdfplumber

st.set_page_config(page_title="ROVO - Debug", page_icon="🔍", layout="wide")
st.title("🔍 Debug PDF")

arquivo = st.file_uploader("Submeter PDF", type=["pdf"])

if arquivo:
    with pdfplumber.open(arquivo) as pdf:
        for i, page in enumerate(pdf.pages):
            st.subheader(f"Página {i+1} — Texto raw")
            st.text(page.extract_text())
