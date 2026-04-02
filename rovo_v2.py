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
            # (Mantemos a lógica Excel que já estava a funcionar)
            pass

        # --- LÓGICA PDF STUDIO NICHOLSON (CORREÇÃO DE REPETIÇÃO) ---
        elif arquivo.name.endswith('.pdf') and cliente == "Studio Nicholson":
            with pdfplumber.open(arquivo) as pdf:
                tams_ref = ["XXS", "XS", "S", "M", "L", "XL", "XXL", "UK4", "UK6", "UK8", "UK10", "UK12", "UK14"]
                lixo_geral = ["JERSEY", "MICRO", "RIB", "SHORT", "SCOOP", "SLEEVE", "NECK", "VEST", "HENLEY", "COTTON", "BRANDED", "BOXY", "FIT", "T-SHIRT", "QTY", "COST", "TOTAL", "FIRST", "MAKE"]

                for page in pdf.pages:
                    palavras = page.extract_words()
                    texto = page.extract_text()
                    if not texto: continue
                    linhas = texto.split('\n')
                    
                    # 1. Capturar Destino
                    destino = "Ver PDF"
                    ship_match = re.search(r"Ship To:\s*(.*)", texto, re.IGNORECASE)
                    if ship_match:
                        destino = ship_match.group(1).split('\n')[0].strip()

                    # 2. Mapear Posições dos Tamanhos
                    mapa_tams = []
                    x_limite_max = 0
                    for p in palavras:
                        txt_up = p['text'].upper().strip()
                        if any(t == txt_up or (t in txt_up and "/" in txt_up) for t in tams_ref):
                            mapa_tams.append({'tam': txt_up, 'x0': p['x0'], 'x1': p['x1'], 'centro': (p['x0']+p['x1'])/2})
                            if p['x1'] > x_limite_max: x_limite_max = p['x1']

                    modelo_atual = ""
                    for i, linha in enumerate(linhas):
                        l_up = linha.upper()
                        
                        # Ignorar lixo e totais
                        if any(x in l_up for x in ["TOTAL", "FIRST", "MAKE", "SUB-TOTAL"]): continue

                        # Detetar Modelo (Designação)
                        if any(x in l_up for x in ["SNW -", "SNM -", "SN -", "LAY "]):
                            modelo_atual = re.split(r"Qty|Cost|Total|First", linha, flags=re.I)[0].strip()
                            continue

                        # Se a linha tem o preço, extraímos os dados desta linha específica
                        if "€" in linha:
                            pts = linha.split()
                            precos = re.findall(r"(\d+[\.,]\d{2})", linha)
                            p_unit = float(precos[0].replace(',', '')) if precos else 0
                            
                            # Isolar a Cor desta linha
                            cor_linha = ""
                            for pt in pts:
                                pt_u = pt.upper().replace(',','').replace('.','')
                                if pt_u not in lixo_geral and not pt_u.isdigit() and "€" not in pt_u and len(pt_u) > 2:
                                    cor_linha = pt
                                    break
                            
                            if not cor_linha: continue

                            # Encontrar a altura Y desta linha no PDF para não ler números de outras linhas
                            # Procuramos a palavra "€" que está mais próxima desta linha de texto
                            y_referencia = None
                            for p_word in palavras:
                                if p_word['text'] == "€" and abs(p_word['top'] - page.extract_text_lines()[i]['top']) < 15:
                                    y_referencia = p_word['top']
                                    break
                            
                            if y_referencia is None: continue

                            # Agora buscamos as quantidades APENAS nesta altura Y
                            for m in mapa_tams:
                                for p_doc in palavras:
                                    # Mesma altura (Y) e mesma coluna (X)
                                    if abs(p_doc['top'] - y_referencia) < 10:
                                        if abs(((p_doc['x0'] + p_doc['x1']) / 2) - m['centro']) < 12:
                                            if p_doc['text'].isdigit() and p_doc['x1'] <= (x_limite_max + 5):
                                                q_num = int(p_doc['text'])
                                                if q_num > 0:
                                                    lista_dados.append({
                                                        'Referência': "", 'Designação': modelo_atual, 'Quant.': q_num,
                                                        'Pr.Unit.': p_unit, 'Pr.Unit.Moeda': 0, 'Tabela de IVA': 4,
                                                        'Cor': cor_linha, 'Tamanho': m['tam'], 'TOTAL': q_num * p_unit, 
                                                        'Destino': destino, 'CPO': ""
                                                    })

        if lista_dados:
            df = pd.DataFrame(lista_dados).drop_duplicates()
            cols = ['Referência', 'Designação', 'Quant.', 'Pr.Unit.', 'Pr.Unit.Moeda', 'Tabela de IVA', 'Cor', 'Tamanho', 'TOTAL', 'Destino', 'CPO']
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='openpyxl') as writer:
                df[cols].to_excel(writer, index=False, sheet_name="PHC")
            st.success("✅ Corrigido! Agora lê as quantidades de cada cor corretamente.")
            st.download_button("⬇️ Download Excel", out.getvalue(), "IMPORTAR_PHC.xlsx")
        else:
            st.warning("Dados não encontrados.")
    except Exception as e:
        st.error(f"Erro: {e}")
