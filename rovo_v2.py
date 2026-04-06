
import streamlit as st
import pandas as pd
import pdfplumber
import re
import io

st.set_page_config(page_title="ROVO - Conversor Universal", page_icon="🚀", layout="wide")
st.sidebar.title("🚀 MENU ROVO")
cliente = st.sidebar.selectbox("Selecione o Cliente", ["Studio Nicholson", "Stussy", "Supreme"])
st.title(f"📦 Conversor: {cliente}")

arquivo = st.file_uploader("Submeter ficheiro", type=["xlsx", "pdf"])

if arquivo:
    try:
        if arquivo.name.endswith(".xlsx"):
            st.info("Lógica Excel ainda não incluída nesta versão de debug.")

        elif arquivo.name.endswith(".pdf") and cliente == "Studio Nicholson":
            with pdfplumber.open(arquivo) as pdf:
                for i, page in enumerate(pdf.pages):
                    st.subheader(f"Página {i+1} — Texto")
                    st.text(page.extract_text())
                    st.subheader(f"Página {i+1} — Palavras")
                    st.write(page.extract_words())

    except Exception as e:
        st.error(f"Erro: {e}")
        st.exception(e)
