import streamlit as st
import pandas as pd
import io
import pdfplumber
import re

st.set_page_config(page_title="ROVO - Conversor Universal", page_icon="🚀", layout="wide")
st.sidebar.title("🚀 MENU ROVO")
cliente = st.sidebar.selectbox("Selecione o Cliente", ["Studio Nicholson", "Stussy", "Supreme"])
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
                        lista_dados.append({'Referência': "", 'Designação': "", 'Quant.': q, 'Pr.Unit.': 0, 'Pr.Unit.Moeda': p, 'Tabela de IVA': 4, 'Cor': row[6], 'Tamanho': row[9], 'TOTAL': q*(p if p else 0), 'Destino': row[4], 'CPO': ""})
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
                                    lista_dados.append({'Referência': "", 'Designação': "", 'Quant.': q, 'Pr.Unit.': 0, 'Pr.Unit.Moeda': p, 'Tabela de IVA': 4, 'Cor': df.iloc[i, 6], 'Tamanho': t_nom, 'TOTAL': q*(p if p else 0), 'Destino': dest, 'CPO': ""})

        # --- LÓGICA PDF STUDIO NICHOLSON (DINÂMICA) ---
        elif arquivo.name.endswith('.pdf') and cliente == "Studio Nicholson":
            with pdfplumber.open(arquivo) as pdf:
                tams_ref = ["XXS", "XS", "S", "M", "L", "XL", "XXL", "UK4", "UK6", "UK8", "UK10", "UK12", "UK14"]
                lixo_cor = ["JERSEY", "MICRO", "RIB", "SHORT", "SCOOP", "SLEEVE", "NECK", "VEST", "HENLEY", "COTTON", "BRANDED", "BOXY", "FIT", "T-SHIRT", "QTY", "COST", "TOTAL"]

                for page in pdf.pages:
                    texto = page.extract_text()
                    if not texto: continue
                    linhas = texto.split('\n')
                    
                    modelo_atual = ""
                    destino_atual = "Ver PDF"
                    mapa_tams = {} 

                    for i, linha in enumerate(linhas):
                        linha_up = linha.upper()
                        
                        if "SHIP TO:" in linha_up and i + 1 < len(linhas):
                            destino_atual = linhas[i+1].strip()
                        
                        if any(x in linha_up for x in ["SNW -", "SNM -", "SN -", "LAY "]):
                            modelo_atual = re.sub(r"\b(QTY|COST|TOTAL|FIRSTMAKE)\b", "", linha, flags=re.I).strip()
                            
                            # Procurar a linha de tamanhos (cabeçalho)
                            for j in range(i, min(i + 4, len(linhas))):
                                partes_h = linhas[j].upper().split()
                                temp_map = {}
                                for idx_p, p in enumerate(partes_h):
                                    for t in tams_ref:
                                        if t == p or (t in p and "/" in p):
                                            temp_map[idx_p] = p
                                if temp_map:
                                    mapa_tams = temp_map
                                    break
                            continue

                        if "€" in linha:
                            partes = linha.split()
                            precos = re.findall(r"(\d+[\.,]\d{2})", linha)
                            p_unit = float(precos[0].replace(',', '')) if precos else 0
                            
                            cor_candidata = ""
                            for p in partes:
                                p_up = p.upper().replace(',', '').replace('.', '')
                                if p_up not in lixo_cor and not p_up.isdigit() and "€" not in p_up and len(p_up) > 2:
                                    cor_candidata = p
                                    break
                            
                            for col_idx, tam_nome in mapa_tams.items():
                                if col_idx < len(partes):
                                    val_q = partes[col_idx]
                                    if val_q.isdigit():
                                        q_num = int(val_q)
                                        if q_num > 0:
                                            lista_dados.append({
                                                'Referência': "", 'Designação': modelo_atual, 'Quant.': q_num,
                                                'Pr.Unit.': p_unit, 'Pr.Unit.Moeda': 0, 'Tabela de IVA': 4,
                                                'Cor': cor_candidata, 'Tamanho': tam_nome,
                                                'TOTAL': q_num * p_unit, 'Destino': destino_atual, 'CPO': ""
                                            })

        if lista_dados:
            df = pd.DataFrame(lista_dados)
            cols = ['Referência', 'Designação', 'Quant.', 'Pr.Unit.', 'Pr.Unit.Moeda', 'Tabela de IVA', 'Cor', 'Tamanho', 'TOTAL', 'Destino', 'CPO']
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='openpyxl') as writer:
                df[cols].to_excel(writer, index=False, sheet_name="PHC")
            st.success("✅ Conversão Concluída!")
            st.download_button("⬇️ Descarregar Excel", out.getvalue(), "IMPORTAR_PHC.xlsx")
    except Exception as e:
        st.error(f"Erro: {e}")
