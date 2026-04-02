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

        # --- LÓGICA STUDIO NICHOLSON (VERSÃO ROBUSTA) ---
        elif arquivo.name.endswith('.pdf') and cliente == "Studio Nicholson":
            with pdfplumber.open(arquivo) as pdf:
                tams_ref = ["XXS", "XS", "S", "M", "L", "XL", "XXL", "UK4", "UK6", "UK8", "UK10", "UK12", "UK14"]
                lixo_geral = ["JERSEY", "MICRO", "RIB", "SHORT", "SCOOP", "SLEEVE", "NECK", "VEST", "HENLEY", "COTTON", "BRANDED", "BOXY", "FIT", "T-SHIRT", "QTY", "COST", "TOTAL", "FIRST/MAKE", "FIRSTMAKE", "DOCKET"]

                for page in pdf.pages:
                    palavras = page.extract_words()
                    texto = page.extract_text()
                    if not texto: continue
                    linhas = texto.split('\n')
                    
                    destino_final = "Ver PDF"
                    ship_match = re.search(r"Ship To:\s*(.*)", texto, re.IGNORECASE)
                    if ship_match:
                        destino_final = ship_match.group(1).split('\n')[0].strip()

                    mapa_posicoes = []
                    for p in palavras:
                        txt = p['text'].upper().strip()
                        if any(t == txt or (t in txt and "/" in txt) for t in tams_ref):
                            mapa_posicoes.append({'tamanho': txt, 'x0': p['x0'], 'x1': p['x1']})

                    modelo_atual = ""
                    for i, linha in enumerate(linhas):
                        linha_up = linha.upper()
                        
                        # Saltar linhas irrelevantes
                        if any(x in linha_up for x in ["TOTAL QTY", "FIRST/MAKE", "DOCKET"]):
                            continue

                        # Identificar Modelo (Designação)
                        if any(x in linha_up for x in ["SNW -", "SNM -", "SN -", "LAY "]):
                            # Limpeza básica para não perder o modelo
                            modelo_atual = linha.split("Qty")[0].split("Cost")[0].strip()
                            continue

                        # Processar linha com preços e quantidades
                        if "€" in linha:
                            partes = linha.split()
                            precos = re.findall(r"(\d+[\.,]\d{2})", linha)
                            p_unit = float(precos[0].replace(',', '')) if precos else 0
                            
                            # Cor
                            cor_candidata = ""
                            for pt in partes:
                                pt_up = pt.upper().replace(',', '').replace('.', '')
                                if pt_up not in lixo_geral and not pt_up.isdigit() and "€" not in pt_up and len(pt_up) > 2:
                                    cor_candidata = pt
                                    break

                            # Procurar números nesta linha usando coordenadas horizontais
                            # Procuramos palavras que estejam na mesma altura (top) aproximada desta linha
                            # y_ref é a altura da primeira palavra da linha (ex: a cor ou o modelo)
                            for m in mapa_posicoes:
                                for p_doc in palavras:
                                    # Se o número está alinhado verticalmente com o tamanho
                                    if abs(p_doc['x0'] - m['x0']) < 25:
                                        # E se o texto desse número aparece nesta linha específica
                                        if p_doc['text'].isdigit() and p_doc['text'] in partes:
                                            # E garantir que não é um número aleatório do topo/fundo
                                            if p_doc['top'] > 150:
                                                q_num = int(p_doc['text'])
                                                if q_num > 0:
                                                    lista_dados.append({
                                                        'Referência': "", 
                                                        'Designação': modelo_atual, 
                                                        'Quant.': q_num, 
                                                        'Pr.Unit.': p_unit, 
                                                        'Pr.Unit.Moeda': 0, 
                                                        'Tabela de IVA': 4, 
                                                        'Cor': cor_candidata, 
                                                        'Tamanho': m['tamanho'], 
                                                        'TOTAL': q_num * p_unit, 
                                                        'Destino': destino_final, 
                                                        'CPO': ""
                                                    })

        if lista_dados:
            df = pd.DataFrame(lista_dados).drop_duplicates()
            cols = ['Referência', 'Designação', 'Quant.', 'Pr.Unit.', 'Pr.Unit.Moeda', 'Tabela de IVA', 'Cor', 'Tamanho', 'TOTAL', 'Destino', 'CPO']
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='openpyxl') as writer:
                df[cols].to_excel(writer, index=False, sheet_name="PHC")
            st.success("✅ Conversão Studio Nicholson restaurada e limpa!")
            st.download_button("⬇️ Descarregar Excel", out.getvalue(), "IMPORTAR_PHC.xlsx")
        else:
            st.warning("Aviso: O ficheiro foi lido mas não foram encontrados dados de quantidades. Verifique se o Cliente selecionado está correto.")

    except Exception as e:
        st.error(f"Erro Crítico: {e}")
