import streamlit as st
import pandas as pd
import io
import pdfplumber
import re

st.set_page_config(page_title="ROVO - Conversor Universal", page_icon="🚀", layout="wide")
st.sidebar.title("🚀 MENU ROVO")
cliente = st.sidebar.selectbox("Selecione o Cliente", ["Stussy", "Supreme", "Studio Nicholson"])
st.title(f"📦 Conversor: {cliente}")

arquivo = st.file_uploader("Submeter ficheiro", type=["xlsx", "pdf"] if cliente == "Studio Nicholson" else ["xlsx"])

if arquivo:
    try:
        lista_dados = []

        if arquivo.name.endswith('.xlsx'):
            xl = pd.ExcelFile(arquivo, engine='openpyxl')
            
            if cliente == "Stussy":
                df = xl.parse(xl.sheet_names[0], header=None)
                # Começamos na linha 2 (índice 1) para saltar o cabeçalho
                for i, row in df.iloc[1:].iterrows():
                    if len(row) >= 14: # Garante que temos até à coluna N
                        # Coluna M (12) = Quantidade
                        q_raw = row[12]
                        # Coluna N (13) = Preço (Pr.Unit.Moeda)
                        p_raw = row[13]
                        
                        q = pd.to_numeric(q_raw, errors='coerce')
                        
                        # Limpeza do Preço (caso venha com texto ou símbolos)
                        if isinstance(p_raw, str):
                            p_clean = re.sub(r'[^\d\.]', '', p_raw.replace(',', '.'))
                            p = pd.to_numeric(p_clean, errors='coerce')
                        else:
                            p = pd.to_numeric(p_raw, errors='coerce')
                        
                        if q and q > 0:
                            lista_dados.append({
                                'Referência': "", 
                                'Designação': row[5] if len(row) > 5 else "", # Coluna F (Estilo)
                                'Quant.': q, 
                                'Pr.Unit.': 0, 
                                'Pr.Unit.Moeda': p if pd.notna(p) else 0, 
                                'Tabela de IVA': 4, 
                                'Cor': row[6] if len(row) > 6 else "", # Coluna G (Color Description)
                                'Tamanho': row[9] if len(row) > 9 else "", # Coluna J (Size)
                                'TOTAL': q * (p if pd.notna(p) else 0), 
                                'Destino': row[4] if len(row) > 4 else "", # Coluna E (Store)
                                'CPO': ""
                            })

            elif cliente == "Supreme":
                # (Lógica Supreme mantida)
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

        elif arquivo.name.endswith('.pdf') and cliente == "Studio Nicholson":
            # (Lógica Studio Nicholson mantida conforme versão anterior estável)
            with pdfplumber.open(arquivo) as pdf:
                tams_ref = ["XXS", "XS", "S", "M", "L", "XL", "XXL", "UK4", "UK6", "UK8", "UK10", "UK12", "UK14"]
                lixo_geral = ["JERSEY", "MICRO", "RIB", "SHORT", "SLEEVE", "NECK", "VEST", "HENLEY", "COTTON", "BRANDED", "BOXY", "FIT", "T-SHIRT", "QTY", "COST", "TOTAL", "FIRST", "MAKE"]
                modelo_atual = ""
                destino_final = "Ver PDF"
                for page in pdf.pages:
                    palavras = page.extract_words()
                    texto = page.extract_text()
                    if not texto: continue
                    linhas = texto.split('\n')
                    ship_match = re.search(r"Ship To:\s*(.*)", texto, re.IGNORECASE)
                    if ship_match:
                        linhas_ship = texto.split("Ship To:")[1].split('\n')
                        destino_final = linhas_ship[1].strip() if len(linhas_ship) > 1 else destino_final
                    mapa_tams = []
                    x_max = 0
                    for p in palavras:
                        t_up = p['text'].upper().strip()
                        if any(t == t_up or (t in t_up and "/" in t_up) for t in tams_ref):
                            mapa_tams.append({'tam': t_up, 'centro': (p['x0']+p['x1'])/2, 'x1': p['x1']})
                            if p['x1'] > x_max: x_max = p['x1']
                    for i, linha in enumerate(linhas):
                        l_up = linha.upper()
                        if any(x in l_up for x in ["TOTAL QTY", "FIRST/MAKE", "SUB-TOTAL"]): continue
                        if any(x in l_up for x in ["SNW -", "SNM -", "SN -", "LAY "]):
                            modelo_atual = re.split(r"Qty|Cost|Total|First", linha, flags=re.I)[0].strip()
                            continue
                        if "€" in linha:
                            pts = linha.split()
                            precos = re.findall(r"(\d+[\.,]\d{2})", linha)
                            p_unit = float(precos[0].replace(',', '')) if precos else 0
                            cor_parts = [pt for pt in pts if pt.upper().replace(',','').replace('.','') not in lixo_geral and not pt.isdigit() and "€" not in pt and len(pt) > 2]
                            cor_final = " ".join(cor_parts[:2])
                            y_ref = next((p_word['top'] for p_word in palavras if p_word['text'] == "€" and abs(p_word['top'] - page.extract_text_lines()[i]['top']) < 20), None)
                            if y_ref is None: continue
                            for m in mapa_tams:
                                for p_doc in palavras:
                                    if abs(p_doc['top'] - y_ref) < 10 and abs(((p_doc['x0'] + p_doc['x1']) / 2) - m['centro']) < 12:
                                        if p_doc['text'].isdigit() and p_doc['x1'] <= (x_max + 10):
                                            q_num = int(p_doc['text'])
                                            if q_num > 0:
                                                lista_dados.append({'Referência': "", 'Designação': modelo_atual, 'Quant.': q_num, 'Pr.Unit.': p_unit, 'Pr.Unit.Moeda': 0, 'Tabela de IVA': 4, 'Cor': cor_final, 'Tamanho': m['tam'], 'TOTAL': q_num * p_unit, 'Destino': destino_final, 'CPO': ""})

        if lista_dados:
            df_final = pd.DataFrame(lista_dados).drop_duplicates()
            cols = ['Referência', 'Designação', 'Quant.', 'Pr.Unit.', 'Pr.Unit.Moeda', 'Tabela de IVA', 'Cor', 'Tamanho', 'TOTAL', 'Destino', 'CPO']
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='openpyxl') as writer:
                df_final[cols].to_excel(writer, index=False, sheet_name="PHC")
            st.success(f"✅ Ficheiro {cliente} processado com sucesso!")
            st.download_button("⬇️ Descarregar Excel", out.getvalue(), f"IMPORTAR_{cliente}.xlsx")
        else:
            st.warning("Não foram encontrados dados no ficheiro.")

    except Exception as e:
        st.error(f"Ocorreu um erro: {e}")
