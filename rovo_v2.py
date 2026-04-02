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

        # --- LÓGICA PDF STUDIO NICHOLSON (MAPEAMENTO POR COORDENADAS) ---
        elif arquivo.name.endswith('.pdf') and cliente == "Studio Nicholson":
            with pdfplumber.open(arquivo) as pdf:
                tams_ref = ["XXS", "XS", "S", "M", "L", "XL", "XXL", "UK4", "UK6", "UK8", "UK10", "UK12", "UK14"]
                lixo_cor = ["JERSEY", "MICRO", "RIB", "SHORT", "SCOOP", "SLEEVE", "NECK", "VEST", "HENLEY", "COTTON", "BRANDED", "BOXY", "FIT", "T-SHIRT", "QTY", "COST", "TOTAL"]

                for page in pdf.pages:
                    # Extraímos as palavras com coordenadas (x0, x1)
                    palavras = page.extract_words()
                    linhas_texto = page.extract_text().split('\n')
                    
                    destino_atual = "Ver PDF"
                    ship_match = re.search(r"Ship To:\s*(.*)", page.extract_text(), re.IGNORECASE)
                    if ship_match:
                        destino_atual = ship_match.group(1).split('\n')[0].strip()

                    mapa_posicoes = [] # Lista de dicionários {tamanho: str, x_centro: float}

                    # 1. Primeiro passamos para identificar os cabeçalhos de tamanhos e suas posições
                    for p in palavras:
                        txt = p['text'].upper().strip()
                        # Detetar tamanhos (incluindo os com barra UK4/IT36)
                        if any(t == txt or (t in txt and "/" in txt) for t in tams_ref):
                            mapa_posicoes.append({
                                'tamanho': txt,
                                'x0': p['x0'],
                                'x1': p['x1'],
                                'centro': (p['x0'] + p['x1']) / 2
                            })

                    # 2. Processar linha a linha para extrair Modelo, Cor e Quantidades
                    modelo_atual = ""
                    for linha in linhas_texto:
                        linha_up = linha.upper()
                        
                        # Identificar Modelo
                        if any(x in linha_up for x in ["SNW -", "SNM -", "SN -", "LAY "]):
                            modelo_atual = re.sub(r"\b(QTY|COST|TOTAL|FIRSTMAKE)\b", "", linha, flags=re.I).strip()
                            continue

                        # Identificar Linha de Dados (Preço e Qtds)
                        if "€" in linha:
                            partes_linha = linha.split()
                            
                            # Preço Unitário
                            precos = re.findall(r"(\d+[\.,]\d{2})", linha)
                            p_unit = float(precos[0].replace(',', '')) if precos else 0
                            
                            # Cor
                            cor_candidata = ""
                            for pt in partes_linha:
                                pt_up = pt.upper().replace(',', '').replace('.', '')
                                if pt_up not in lixo_cor and not pt_up.isdigit() and "€" not in pt_up and len(pt_up) > 2:
                                    cor_candidata = pt
                                    break

                            # Extrair palavras da linha atual com coordenadas para bater com o cabeçalho
                            palavras_linha = [pw for pw in palavras if pw['top'] > 0 and abs(pw['top'] - page.extract_text_lines()[0]['top']) < 500] # Simplificação de busca
                            # Versão robusta: re-extraímos as palavras desta linha específica
                            
                            # Para cada tamanho mapeado no cabeçalho, procuramos um número abaixo dele
                            for m in mapa_posicoes:
                                # Procuramos na linha de texto atual por números que estejam "perto" da coordenada X do tamanho
                                # Vamos usar uma margem de erro de 15 pixels
                                for p_doc in palavras:
                                    # Se a palavra está na mesma altura (Y) que a linha do € e na mesma largura (X) que o cabeçalho
                                    if abs(p_doc['x0'] - m['x0']) < 20 and p_doc['text'].isdigit():
                                        q_num = int(p_doc['text'])
                                        if q_num > 0 and p_doc['top'] > 100: # Evita apanhar números do topo da página
                                            # Verificar se este número pertence à linha atual (pelo texto)
                                            if p_doc['text'] in partes_linha:
                                                lista_dados.append({
                                                    'Referência': "", 'Designação': modelo_atual, 'Quant.': q_num,
                                                    'Pr.Unit.': p_unit, 'Pr.Unit.Moeda': 0, 'Tabela de IVA': 4,
                                                    'Cor': cor_candidata, 'Tamanho': m['tamanho'],
                                                    'TOTAL': q_num * p_unit, 'Destino': destino_atual, 'CPO': ""
                                                })

        if lista_dados:
            # Remover duplicados que possam surgir da busca por coordenadas
            df = pd.DataFrame(lista_dados).drop_duplicates()
            cols = ['Referência', 'Designação', 'Quant.', 'Pr.Unit.', 'Pr.Unit.Moeda', 'Tabela de IVA', 'Cor', 'Tamanho', 'TOTAL', 'Destino', 'CPO']
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='openpyxl') as writer:
                df[cols].to_excel(writer, index=False, sheet_name="PHC")
            st.success("✅ Conversão por Coordenadas Concluída!")
            st.download_button("⬇️ Descarregar Excel", out.getvalue(), "IMPORTAR_PHC.xlsx")
    except Exception as e:
        st.error(f"Erro: {e}")
