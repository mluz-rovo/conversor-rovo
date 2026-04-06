import streamlit as st
import pandas as pd
import io
import pdfplumber
import re

st.set_page_config(page_title="ROVO - Conversor Universal", page_icon="🚀", layout="wide")

st.sidebar.title("🚀 MENU ROVO")
cliente = st.sidebar.selectbox("Selecione o Cliente", ["Stussy", "Supreme", "Studio Nicholson"])

# --- NOVOS CAMPOS PARA SUPREME ---
ref_manual = ""
des_manual = ""
if cliente == "Supreme":
    st.sidebar.write("---")
    st.sidebar.subheader("📝 Dados do PO Supreme")
    ref_manual = st.sidebar.text_input("Referência (ex: FW24)")
    des_manual = st.sidebar.text_input("Designação (ex: Box Logo Hooded)")

st.title(f"📦 Conversor: {cliente}")

arquivo = st.file_uploader("Submeter ficheiro", type=["xlsx", "pdf"] if cliente == "Studio Nicholson" else ["xlsx"])

if arquivo:
    try:
        lista_dados = []

        if arquivo.name.endswith('.xlsx'):
            xl = pd.ExcelFile(arquivo, engine='openpyxl')
            
            # --- LÓGICA STUSSY ---
            if cliente == "Stussy":
                df = xl.parse(xl.sheet_names[0], header=None)
                for i, row in df.iloc[1:].iterrows():
                    if len(row) >= 14:
                        q = pd.to_numeric(row[12], errors='coerce')
                        p_raw = row[13]
                        if isinstance(p_raw, str):
                            p = pd.to_numeric(re.sub(r'[^\d\.]', '', p_raw.replace(',', '.')), errors='coerce')
                        else:
                            p = pd.to_numeric(p_raw, errors='coerce')
                        
                        if q and q > 0:
                            lista_dados.append({
                                'Referência': "", 'Designação': row[8] if len(row) > 8 else "", 
                                'Quant.': q, 'Pr.Unit.': 0, 'Pr.Unit.Moeda': p if pd.notna(p) else 0, 
                                'Tabela de IVA': 4, 'Cor': row[7] if len(row) > 7 else "", 
                                'Tamanho': row[9] if len(row) > 9 else "", 'TOTAL': q * (p if pd.notna(p) else 0), 
                                'Destino': row[4] if len(row) > 4 else "Geral", 'CPO': ""
                            })

            # --- LÓGICA SUPREME (COM INPUTS MANUAIS) ---
            elif cliente == "Supreme":
                for aba in xl.sheet_names:
                    if "TOTAL" in aba.upper(): continue
                    df = xl.parse(aba, header=None)
                    tams = {c: str(df.iloc[14, c]) for c in range(9, 16) if pd.notna(df.iloc[14, c])}
                    
                    for start in range(16, len(df), 14):
                        dest = str(df.iloc[start, 0]).strip()
                        if not dest or dest == "nan": dest = "Indefinido"
                        
                        for i in range(start + 1, start + 13):
                            if i >= len(df) or pd.isna(df.iloc[i, 6]): continue
                            p = pd.to_numeric(df.iloc[i, 17], errors='coerce')
                            for c_idx, t_nom in tams.items():
                                q = pd.to_numeric(df.iloc[i, c_idx], errors='coerce')
                                if q and q > 0:
                                    lista_dados.append({
                                        'Referência': ref_manual, # USA O CAMPO MANUAL
                                        'Designação': des_manual, # USA O CAMPO MANUAL
                                        'Quant.': q, 'Pr.Unit.': 0, 'Pr.Unit.Moeda': p if pd.notna(p) else 0, 
                                        'Tabela de IVA': 4, 'Cor': df.iloc[i, 6], 'Tamanho': t_nom, 
                                        'TOTAL': q * (p if pd.notna(p) else 0), 'Destino': dest, 'CPO': ""
                                    })

        # --- LÓGICA STUDIO NICHOLSON ---
        elif arquivo.name.endswith('.pdf') and cliente == "Studio Nicholson":
            with pdfplumber.open(arquivo) as pdf:
                tams_ref = ["XXS", "XS", "S", "M", "L", "XL", "XXL", "
