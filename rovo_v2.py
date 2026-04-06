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
        
        # --- LÓGICA EXCEL MANTIDA ---
        if arquivo.name.endswith('.xlsx'):
            pass # (Lógica Stussy/Supreme omitida aqui para brevidade, mas mantida no teu ficheiro)

        # --- LÓGICA STUDIO NICHOLSON (AFINAÇÃO FINAL) ---
        elif arquivo.name.endswith('.pdf') and cliente == "Studio Nicholson":
            with pdfplumber.open(arquivo) as pdf:
                tams_ref = ["XXS", "XS", "S", "M", "L", "XL", "XXL", "UK4", "UK6", "UK8", "UK10", "UK12", "UK14"]
                lixo_geral = ["JERSEY", "MICRO", "RIB", "SHORT", "SLEEVE", "NECK", "VEST", "HENLEY", "COTTON", "BRANDED", "BOXY", "FIT", "T-SHIRT", "QTY", "COST", "TOTAL", "FIRST", "MAKE", "DOCKET"]

                modelo_atual = ""
                destino_final = "Ver PDF"

                for page in pdf.pages:
                    palavras = page.extract_words()
                    texto = page.extract_text()
                    if not texto: continue
                    linhas = texto.split('\n')
                    
                    # 1. Tentar capturar Destino na página
                    ship_match = re.search(r"Ship To:\s*(.*)", texto, re.IGNORECASE)
                    if ship_match:
                        linhas_ship = texto.split("Ship To:")[1].split('\n')
                        destino_final = linhas_ship[1].strip() if len(linhas_ship) > 1 else destino_final

                    # 2. Mapear Posições dos Tamanhos
                    mapa_tams = []
                    x_limite_max = 0
                    for p in palavras:
                        txt_up = p['text'].upper().strip()
                        if any(t == txt_up or (t in txt_up and "/" in txt_up) for t in tams_ref):
                            mapa_tams.append({'tam': txt_up, 'centro': (p['x0']+p['x1'])/2, 'x1': p['x1']})
                            if p['x1'] > x_limite_max: x_limite_max = p['x1']

                    for i, linha in enumerate(linhas):
                        l_up = linha.upper()
                        
                        # Ignorar lixo e totais
                        if any(x in l_up for x in ["TOTAL QTY", "FIRST/MAKE", "SUB-TOTAL"]): continue

                        # 3. Detetar e MANTER o Modelo (Designação)
                        if any(x in l_up for x in ["SNW -", "SNM -", "SN -", "LAY "]):
                            modelo_atual = re.split(r"Qty|Cost|Total|First", linha, flags=re.I)[0].strip()
                            continue

                        # 4. Se a linha tem o preço (€), processamos cor e quantidades
                        if "€" in linha:
                            pts = linha.split()
                            precos = re.findall(r"(\d+[\.,]\d{2})", linha)
                            p_unit = float(precos[0].replace(',', '')) if precos else 0
                            
                            # Isolar a Cor (Pega as palavras que não são lixo nem números)
                            cor_parts = []
                            for pt in pts:
                                pt_u = pt.upper().replace(',','').replace('.','')
                                if pt_u not in lixo_geral and not pt_u.isdigit() and "€" not in pt_u and len(pt_u) > 2:
                                    cor_parts.append(pt)
                            
                            cor_final = " ".join(cor_parts[:2]) # Captura "Optic White" (as duas primeiras)

                            # Encontrar altura Y da linha atual
                            y_ref = None
                            for p_word in palavras:
                                if p_word['text'] == "€" and abs(p_word['top'] - page.extract_text_lines()[i]['top']) < 20:
                                    y_ref = p_word['top']
                                    break
                            
                            if y_ref is None: continue

                            # Extrair quantidades baseadas no mapa de tamanhos
                            for m in mapa_tams:
                                for p_doc in palavras:
                                    if abs(p_doc['top'] - y_ref) < 10: # Mesma linha
                                        if abs(((p_doc['x0'] + p_doc['x1']) / 2) - m['centro']) < 12: # Mesma coluna
                                            if p_doc['text'].isdigit() and p_doc['x1'] <= (x_limite_max + 10):
                                                q_num = int(p_doc['text'])
                                                if q_num > 0:
                                                    lista_dados.append({
                                                        'Referência': "", 
                                                        'Designação': modelo_atual, 
                                                        'Quant.': q_num,
                                                        'Pr.Unit.': p_unit, 
                                                        'Pr.Unit.Moeda': 0, 
                                                        'Tabela de IVA': 4,
                                                        'Cor': cor_final, 
                                                        'Tamanho': m['tam'], 
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
            st.success("✅ Tudo corrigido: Cor completa, Designação e Destino ativos!")
            st.download_button("⬇️ Download Excel", out.getvalue(), "IMPORTAR_PHC.xlsx")
    except Exception as e:
        st.error(f"Erro: {e}")
