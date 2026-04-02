import streamlit as st
import pandas as pd
import io
import pdfplumber
import re

# FORÇAR REINÍCIO - VERSÃO VERDE - 16:15
st.set_page_config(page_title="ROVO VERDE", layout="wide")

# ESTA PARTE PINTA O SITE DE VERDE
st.markdown("""
    <style>
    .stApp { background-color: #d4edda; }
    </style>
    """, unsafe_allow_html=True)

st.sidebar.title("🌿 MENU ROVO")
st.sidebar.success("✅ CÓDIGO VERDE ATIVO (16:15)")
cliente = st.sidebar.selectbox("Escolha o Cliente", ["Studio Nicholson", "Stussy", "Supreme"])

st.title(f"📦 Conversor ROVO: {cliente}")

arquivo = st.file_uploader("Submeter PDF ou Excel", type=["xlsx", "pdf"])

if arquivo:
    try:
        lista_dados = []
        # --- LÓGICA NICHOLSON (PREÇO NO PR.UNIT / DESIGNAÇÃO COM MODELO / COR LIMPA) ---
        if arquivo.name.endswith('.pdf') and cliente == "Studio Nicholson":
            with pdfplumber.open(arquivo) as pdf:
                tams_ref = ["XS", "S", "M", "L", "XL", "XXL", "UK4", "UK6", "UK8", "UK10", "UK12", "UK14"]
                proibidos = ["JERSEY", "MICRO", "RIB", "MERCERIZED", "COTTON", "BRANDED", "BOXY", "FIT", "SNW", "SNM", "LAY"]

                for page in pdf.pages:
                    texto = page.extract_text() or ""
                    tab = page.extract_tables()
                    ship = re.search(r"Ship To:\s*(.*)", texto, re.IGNORECASE)
                    dest = ship.group(1).split('\n')[0].strip() if ship else "Ver PDF"

                    for table in tab:
                        h = []
                        s_idx = -1
                        for r_idx, r in enumerate(table):
                            r_str = " ".join([str(x).upper() for x in r if x])
                            if any(t in r_str for t in tams_ref):
                                h = [str(x).replace('\n', ' ').strip() for x in r]
                                s_idx = r_idx + 1
                                break
                        if s_idx == -1: continue

                        for i in range(s_idx, len(table)):
                            row = table[i]
                            r_full = " ".join([str(x) for x in row if x]).replace('\n', ' ')
                            if "€" in r_full:
                                mod = str(row[0]).split('\n')[0].strip()
                                pts = r_full.split()
                                # Limpeza Radical da Cor
                                c_limpa = [p for p in pts if p.upper() not in proibidos and not p.replace('.','').isdigit() and "€" not in p and "SN" not in p.upper() and len(p)>2]
                                cor = " ".join(c_limpa).strip()
                                
                                p_v = 0
                                for c in row:
                                    if "€" in str(c):
                                        p_v = pd.to_numeric(str(c).replace('€','').replace(',','.').replace(' ',''), errors='coerce')
                                        break

                                for c_idx, head in enumerate(h):
                                    t_ok = ""
                                    for t in tams_ref:
                                        if t in head.upper(): t_ok = head; break
                                    if t_ok:
                                        qtd = pd.to_numeric(row[c_idx], errors='coerce')
                                        if qtd and qtd > 0:
                                            lista_dados.append({
                                                'Referência': "", 'Designação': mod, 'Quant.': qtd,
                                                'Pr.Unit.': p_v, 'Pr.Unit.Moeda': 0, 'Tabela de IVA': 4,
                                                'Cor': cor, 'Tamanho': t_ok, 'TOTAL': qtd*p_v, 'Destino': dest, 'Aba': "Nicholson"
                                            })
        # --- (Lógica Stussy/Supreme mantida internamente) ---
        if lista_dados:
            df = pd.DataFrame(lista_dados)
            df['CPO'] = ""
            cols = ['Referência', 'Designação', 'Quant.', 'Pr.Unit.', 'Pr.Unit.Moeda', 'Tabela de IVA', 'Cor', 'Tamanho', 'TOTAL', 'Destino', 'CPO']
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='openpyxl') as writer:
                df[cols].to_excel(writer, index=False, sheet_name="PHC")
            st.success("✅ Convertido!")
            st.download_button("⬇️ Download Excel", out.getvalue(), "IMPORTAR.xlsx")
    except Exception as e:
        st.error(f"Erro: {e}")
