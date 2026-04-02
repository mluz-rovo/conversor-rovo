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

        # --- LÓGICA STUDIO NICHOLSON (V12 - ESTABILIZADA) ---
        elif arquivo.name.endswith('.pdf') and cliente == "Studio Nicholson":
            with pdfplumber.open(arquivo) as pdf:
                tams_ref = ["XXS", "XS", "S", "M", "L", "XL", "XXL", "UK4", "UK6", "UK8", "UK10", "UK12", "UK14"]
                lixo_geral = ["JERSEY", "MICRO", "RIB", "SHORT", "SCOOP", "SLEEVE", "NECK", "VEST", "HENLEY", "COTTON", "BRANDED", "BOXY", "FIT", "T-SHIRT", "QTY", "COST", "TOTAL", "FIRST", "MAKE", "SAMPLE"]

                for page in pdf.pages:
                    palavras = page.extract_words()
                    texto = page.extract_text()
                    if not texto: continue
                    linhas = texto.split('\n')
                    
                    destino = "Ver PDF"
                    ship_match = re.search(r"Ship To:\s*(.*)", texto, re.IGNORECASE)
                    if ship_match:
                        destino = ship_match.group(1).split('\n')[0].strip()

                    # 1. Mapear Cabeçalhos e Limite Direito
                    mapa_tams = []
                    x_limite = 0
                    for p in palavras:
                        txt_up = p['text'].upper().strip()
                        if any(t == txt_up or (t in txt_up and "/" in txt_up) for t in tams_ref):
                            mapa_tams.append({'tam': txt_up, 'centro_x': (p['x0'] + p['x1']) / 2})
                            if p['x1'] > x_limite: x_limite = p['x1']

                    modelo_atual = ""
                    for i, linha in enumerate(linhas):
                        l_up = linha.upper()
                        
                        # Saltar lixo
                        if any(x in l_up for x in ["TOTAL QTY", "FIRST/MAKE", "FIRST MAKE", "SUB-TOTAL"]): continue

                        # Capturar Modelo
                        if any(x in l_up for x in ["SNW -", "SNM -", "SN -", "LAY "]):
                            modelo_atual = re.split(r"Qty|Cost|Total|First", linha, flags=re.I)[0].strip()
                            continue

                        # Se a linha tem o preço, processamos as quantidades
                        if "€" in linha:
                            precos = re.findall(r"(\d+[\.,]\d{2})", linha)
                            p_unit = float(precos[0].replace(',', '')) if precos else 0
                            
                            # Isolar a Cor
                            cor = "Ver PDF"
                            for pt in linha.split():
                                pt_u = pt.upper().replace(',','').replace('.','')
                                if pt_u not in lixo_geral and not pt_u.isdigit() and "€" not in pt_u and len(pt_u) > 2:
                                    cor = pt
                                    break
                            
                            # Coordenada Y da linha atual (baseada no símbolo €)
                            euro_word = [p for p in palavras if p['text'] == "€" and abs(p['top'] - page.extract_text_lines()[i]['top']) < 50]
                            if not euro_word: continue
                            y_linha = euro_word[0]['top']

                            for m in mapa_tams:
                                for p_doc in palavras:
                                    # Critério 1: Mesma altura horizontal (Y)
                                    # Critério 2: Mesmo centro vertical (X)
                                    # Critério 3: É um número
                                    if abs(p_doc['top'] - y_linha) < 8:
                                        if abs(((p_doc['x0'] + p_doc['x1']) / 2) - m['centro_x']) < 10:
                                            if p_doc['text'].isdigit() and p_doc['x1'] <= (x_limite + 5):
                                                q_num = int(p_doc['text'])
                                                if q_num > 0:
                                                    lista_dados.append({
                                                        'Referência': "", 'Designação': modelo_atual, 'Quant.': q_num,
                                                        'Pr.Unit.': p_unit, 'Pr.Unit.Moeda': 0, 'Tabela de IVA': 4,
                                                        'Cor': cor, 'Tamanho': m['tam'], 'TOTAL': q_num * p_unit, 
                                                        'Destino': destino, 'CPO': ""
                                                    })

        if lista_dados:
            df = pd.DataFrame(lista_dados).drop_duplicates()
            # Filtro Final de Segurança
            df = df[~df['Tamanho'].str.contains("FIRST|MAKE|TOTAL|QTY", case=False, na=False)]
            
            cols = ['Referência', 'Designação', 'Quant.', 'Pr.Unit.', 'Pr.Unit.Moeda', 'Tabela de IVA', 'Cor', 'Tamanho', 'TOTAL', 'Destino', 'CPO']
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='openpyxl') as writer:
                df[cols].to_excel(writer, index=False, sheet_name="PHC")
            st.success("✅ V12 Estável: Quantidades e Nomes limpos!")
            st.download_button("⬇️ Download Excel", out.getvalue(), "IMPORTAR_PHC.xlsx")
        else:
            st.warning("Dados não encontrados.")
    except Exception as e:
        st.error(f"Erro: {e}")
