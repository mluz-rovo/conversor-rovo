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

        # --- LÓGICA STUDIO NICHOLSON (LIMITADOR DE COLUNAS) ---
        elif arquivo.name.endswith('.pdf') and cliente == "Studio Nicholson":
            with pdfplumber.open(arquivo) as pdf:
                tams_ref = ["XXS", "XS", "S", "M", "L", "XL", "XXL", "UK4", "UK6", "UK8", "UK10", "UK12", "UK14"]
                lixo_geral = ["JERSEY", "MICRO", "RIB", "SHORT", "SCOOP", "SLEEVE", "NECK", "VEST", "HENLEY", "COTTON", "BRANDED", "BOXY", "FIT", "T-SHIRT", "QTY", "COST"]

                for page in pdf.pages:
                    texto = page.extract_text()
                    if not texto: continue
                    linhas = texto.split('\n')
                    palavras_pdf = page.extract_words()
                    
                    destino = "Ver PDF"
                    for i, l in enumerate(linhas):
                        if "Ship To:" in l and i+1 < len(linhas):
                            destino = linhas[i+1].strip()
                            break

                    # 1. Identificar Posições dos Tamanhos e calcular o LIMITE DIREITO
                    mapa = []
                    x_limite_direito = 0
                    for p in palavras_pdf:
                        t_up = p['text'].upper().strip()
                        if any(t == t_up or (t in t_up and "/" in t_up) for t in tams_ref):
                            mapa.append({'tam': t_up, 'x0': p['x0'], 'x1': p['x1']})
                            if p['x1'] > x_limite_direito:
                                x_limite_direito = p['x1'] # O tamanho mais à direita define o limite

                    modelo_atual = ""
                    for linha in linhas:
                        l_up = linha.upper()
                        
                        # Filtro de Totais e First Make
                        if any(x in l_up for x in ["TOTAL", "FIRST", "MAKE", "SAMPLE", "DOCKET"]):
                            continue

                        # Capturar Modelo
                        if any(x in l_up for x in ["SNW -", "SNM -", "SN -", "LAY "]):
                            modelo_atual = re.split(r"Qty|Cost|Total|First", linha, flags=re.I)[0].strip()
                            continue

                        if "€" in linha:
                            pts = linha.split()
                            precos = re.findall(r"(\d+[\.,]\d{2})", linha)
                            p_val = float(precos[0].replace(',', '')) if precos else 0
                            
                            # Cor
                            cor_candidata = ""
                            for pt in pts:
                                pt_u = pt.upper().replace(',','').replace('.','')
                                if pt_u not in lixo_geral and not pt_u.isdigit() and "€" not in pt_u and len(pt_u) > 2:
                                    cor_candidata = pt
                                    break
                            
                            if not cor_candidata: continue

                            # Capturar Quantidades (Com trava de segurança no limite X)
                            for m in mapa:
                                for p_doc in palavras_pdf:
                                    # Só aceita o número se estiver alinhado com o tamanho E à esquerda do limite
                                    if abs(p_doc['x0'] - m['x0']) < 20 and p_doc['text'].isdigit() and p_doc['text'] in pts:
                                        # GARANTIA: O número tem de estar dentro da zona de tamanhos
                                        if p_doc['x1'] <= (x_limite_direito + 15): 
                                            q_num = int(p_doc['text'])
                                            if q_num > 0 and p_doc['top'] > 120:
                                                lista_dados.append({
                                                    'Referência': "", 'Designação': modelo_atual, 'Quant.': q_num,
                                                    'Pr.Unit.': p_val, 'Pr.Unit.Moeda': 0, 'Tabela de IVA': 4,
                                                    'Cor': cor_candidata, 'Tamanho': m['tam'], 'TOTAL': q_num * p_val, 
                                                    'Destino': destino, 'CPO': ""
                                                })

        if lista_dados:
            df = pd.DataFrame(lista_dados).drop_duplicates()
            cols = ['Referência', 'Designação', 'Quant.', 'Pr.Unit.', 'Pr.Unit.Moeda', 'Tabela de IVA', 'Cor', 'Tamanho', 'TOTAL', 'Destino', 'CPO']
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='openpyxl') as writer:
                df[cols].to_excel(writer, index=False, sheet_name="PHC")
            st.success("✅ Conversão Concluída! Totais à direita ignorados com sucesso.")
            st.download_button("⬇️ Descarregar Excel", out.getvalue(), "IMPORTAR_PHC.xlsx")
    except Exception as e:
        st.error(f"Erro: {e}")
