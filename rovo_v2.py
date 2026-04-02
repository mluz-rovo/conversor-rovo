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

        # --- LÓGICA PDF STUDIO NICHOLSON (ROBUSTA) ---
        elif arquivo.name.endswith('.pdf') and cliente == "Studio Nicholson":
            with pdfplumber.open(arquivo) as pdf:
                tams_ref = ["XS", "S", "M", "L", "XL", "XXL", "UK4", "UK6", "UK8", "UK10", "UK12", "UK14"]
                palavras_lixo = ["JERSEY", "MICRO", "RIB", "MERCERIZED", "COTTON", "BRANDED", "BOXY", "FIT", "T-SHIRT", "QTY", "OTY", "TOTAL", "COST", "FIRSTMAKE"]

                for page in pdf.pages:
                    linhas = page.extract_text().split('\n')
                    modelo_atual = ""
                    destino = "Ver PDF"
                    
                    for idx, linha in enumerate(linhas):
                        linha_up = linha.upper()
                        
                        # 1. Encontrar Destino
                        if "SHIP TO:" in linha_up:
                            destino = linhas[idx+1].strip() if idx+1 < len(linhas) else "Ver PDF"
                        
                        # 2. Encontrar Modelo (Designação)
                        if any(x in linha_up for x in ["SNW-", "SNM-", "LAY "]):
                            modelo_atual = linha.strip()
                            continue

                        # 3. Encontrar Linha de Dados (Preço e Quantidades)
                        if "€" in linha:
                            partes = linha.split()
                            
                            # Extrair Preço Unitário
                            precos = [p for p in partes if "€" in p or p.replace(',','').replace('.','').isdigit()]
                            p_unit = 0
                            for p in partes:
                                if "€" in p:
                                    idx_p = partes.index(p)
                                    val_txt = partes[idx_p+1] if idx_p+1 < len(partes) else p
                                    p_unit = pd.to_numeric(val_txt.replace('€','').replace(',',''), errors='coerce')
                                    break
                            
                            # Extrair Cor (Primeiras palavras da linha que não sejam lixo ou números)
                            cor_partes = []
                            for p in partes:
                                p_clean = p.upper().replace(',','')
                                if p_clean not in palavras_lixo and not p_clean.replace('.','').isdigit() and "€" not in p_clean and len(p_clean) > 2:
                                    cor_partes.append(p)
                            cor_final = " ".join(cor_partes[:2]) # Pega as duas primeiras palavras da cor
                            
                            # Extrair Quantidades
                            numeros = [n for n in partes if n.replace('.','').isdigit() and "€" not in n]
                            # Geralmente o último número é o total da linha, removemos
                            qts = numeros[:-1] if len(numeros) > 1 else numeros

                            for i_q, val_q in enumerate(qts):
                                if i_q < len(tams_ref):
                                    q_num = pd.to_numeric(val_q, errors='coerce')
                                    if q_num and q_num > 0:
                                        lista_dados.append({
                                            'Referência': "", 
                                            'Designação': modelo_atual, 
                                            'Quant.': q_num, 
                                            'Pr.Unit.': p_unit, 
                                            'Pr.Unit.Moeda': 0, 
                                            'Tabela de IVA': 4, 
                                            'Cor': cor_final, 
                                            'Tamanho': tams_ref[i_q], 
                                            'TOTAL': q_num * p_unit, 
                                            'Destino': destino, 
                                            'CPO': ""
                                        })

        if lista_dados:
            df_final = pd.DataFrame(lista_dados)
            cols = ['Referência', 'Designação', 'Quant.', 'Pr.Unit.', 'Pr.Unit.Moeda', 'Tabela de IVA', 'Cor', 'Tamanho', 'TOTAL', 'Destino', 'CPO']
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='openpyxl') as writer:
                df_final[cols].to_excel(writer, index=False, sheet_name="PHC")
            st.success("✅ Conversão concluída!")
            st.download_button("⬇️ Descarregar Excel", out.getvalue(), "IMPORTAR.xlsx")
        else:
            st.warning("Não foram encontrados dados. Verifique se o Cliente selecionado corresponde ao ficheiro.")

    except Exception as e:
        st.error(f"Erro: {e}")
