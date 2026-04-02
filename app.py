import streamlit as st
import pandas as pd
import io
import pdfplumber
import re

# FORÇAR REINÍCIO DO SERVIDOR - 15:55
st.set_page_config(page_title="ROVO - CONVERSOR V3", layout="wide")

# Mudar a cor para sabermos que o código atualizou
st.markdown("""
    <style>
    .stApp { background-color: #f0f2f6; }
    </style>
    """, unsafe_allow_html=True)

st.sidebar.title("MENU ROVO")
cliente = st.sidebar.selectbox("Cliente", ["Stussy", "Supreme", "Studio Nicholson"])
st.sidebar.success("🚀 VERSÃO V3 ATIVA (15:55)")

st.title(f"📦 Conversor: {cliente}")

arquivo = st.file_uploader("Submeter ficheiro", type=["xlsx", "pdf"])

if arquivo:
    try:
        lista_dados = []

        # --- LÓGICA EXCEL (STUSSY / SUPREME) ---
        if arquivo.name.endswith('.xlsx'):
            xl = pd.ExcelFile(arquivo)
            if cliente == "Stussy":
                df = xl.parse(xl.sheet_names[0], header=None)
                for i, row in df.iloc[1:].iterrows():
                    q, p = pd.to_numeric(row[12], errors='coerce'), pd.to_numeric(row[17], errors='coerce')
                    if q and q > 0:
                        lista_dados.append({'Referência': "", 'Designação': "", 'Quant.': q, 'Pr.Unit.': 0, 'Pr.Unit.Moeda': p, 'Tabela de IVA': 4, 'Cor': row[6], 'Tamanho': row[9], 'TOTAL': q*(p if p else 0), 'Destino': row[4], 'Aba': "Stussy"})

            elif cliente == "Supreme":
                for aba in xl.sheet_names:
                    if "TOTAL" in aba.upper(): continue
                    df = xl.parse(aba, header=None)
                    tams = {c: str(df.iloc[14, c]) for c in range(9, 16) if pd.notna(df.iloc[14, c])}
                    for start in range(16, len(df), 14):
                        dest = str(df.iloc[start, 0]).strip()
                        for i in range(start + 1, start + 13):
                            if i >= len(df) or pd.isna(df.iloc[i, 6]): continue
                            p = pd.to_numeric(df.iloc[i, 17], errors='coerce')
                            for c_idx, t_nom in tams.items():
                                q = pd.to_numeric(df.iloc[i, c_idx], errors='coerce')
                                if q and q > 0:
                                    lista_dados.append({'Referência': "", 'Designação': "", 'Quant.': q, 'Pr.Unit.': 0, 'Pr.Unit.Moeda': p, 'Tabela de IVA': 4, 'Cor': df.iloc[i, 6], 'Tamanho': t_nom, 'TOTAL': q*(p if p else 0), 'Destino': dest, 'Aba': aba})

        # --- LÓGICA NICHOLSON (CORREÇÃO DE COLUNAS E COR) ---
        elif arquivo.name.endswith('.pdf') and cliente == "Studio Nicholson":
            with pdfplumber.open(arquivo) as pdf:
                tams_ref = ["XS", "S", "M", "L", "XL", "XXL", "UK4", "UK6", "UK8", "UK10", "UK12", "UK14"]
                proibidos = ["JERSEY", "MICRO", "RIB", "MERCERIZED", "COTTON", "BRANDED", "BOXY", "FIT", "SNW", "SNM", "LAY"]

                for page in pdf.pages:
                    texto = page.extract_text() or ""
                    tab = page.extract_tables()
                    dest = "Ver PDF"
                    ship = re.search(r"Ship To:\s*(.*)", texto, re.IGNORECASE)
                    if ship: dest = ship.group(1).split('\n')[0].strip()

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
                                # Filtro de Cor
                                pts = r_full.split()
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
                                        qtd = pd.to_numeric(row[col_idx], errors='coerce') if 'col_idx' in locals() else pd.to_numeric(row[c_idx], errors='coerce')
                                        if qtd and qtd > 0:
                                            lista_dados.append({
                                                'Referência': "", 'Designação': mod, 'Quant.': qtd,
                                                'Pr.Unit.': p_v, 'Pr.Unit.Moeda': 0, 'Tabela de IVA': 4,
                                                'Cor': cor, 'Tamanho': t_ok, 'TOTAL': qtd*p_v, 'Destino': dest, 'Aba': "Nicholson"
                                            })

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
