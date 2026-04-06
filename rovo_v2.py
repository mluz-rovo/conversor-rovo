import streamlit as st
import pandas as pd
import io
import pdfplumber
import re

st.set_page_config(page_title="ROVO - Conversor Universal", page_icon="🚀", layout="wide")

st.sidebar.title("🚀 MENU ROVO")
cliente = st.sidebar.selectbox("Selecione o Cliente", ["Stussy", "Supreme", "Studio Nicholson"])

# --- CAMPOS MANUAIS APENAS PARA SUPREME ---
ref_manual = ""
des_manual = ""

if cliente == "Supreme":
    st.sidebar.write("---")
    st.sidebar.subheader("📝 Dados Fixos Supreme")
    ref_manual = st.sidebar.text_input("Referência para PHC", placeholder="Ex: FW24-001")
    des_manual = st.sidebar.text_input("Designação para PHC", placeholder="Ex: Box Logo Hooded")
    st.sidebar.caption("Estes valores aparecerão em todas as linhas.")

st.title(f"📦 Conversor: {cliente}")

arquivo = st.file_uploader("Submeter ficheiro", type=["xlsx", "pdf"] if cliente == "Studio Nicholson" else ["xlsx"])

if arquivo:
    try:
        lista_dados = []

        if arquivo.name.endswith('.xlsx'):
            xl = pd.ExcelFile(arquivo, engine='openpyxl')
            
            # --- LÓGICA STUSSY ---
            if cliente == "Stussy":
                df = xl.parse(xl.sheet_names[0], header=None)
                for i, row in df.iloc[1:].iterrows():
                    if len(row) >= 14:
                        q = pd.to_numeric(row[12], errors='coerce')
                        p_raw = row[13]
                        if isinstance(p_raw, str):
                            p = pd.to_numeric(re.sub(r'[^\d\.]', '', p_raw.replace(',', '.')), errors='coerce')
                        else:
                            p = pd.to_numeric(p_raw, errors='coerce')
                        
                        if q and q > 0:
                            lista_dados.append({
                                'Referência': "", 
                                'Designação': row[8] if len(row) > 8 else "", 
                                'Quant.': q, 
                                'Pr.Unit.': 0, 
                                'Pr.Unit.Moeda': p if pd.notna(p) else 0, 
                                'Tabela de IVA': 4, 
                                'Cor': row[7] if len(row) > 7 else "", 
                                'Tamanho': row[9] if len(row) > 9 else "", 
                                'TOTAL': q * (p if pd.notna(p) else 0), 
                                'Destino': row[4] if len(row) > 4 else "Geral", 
                                'Nr. CPO': "",
                                'Nr. SPO': "",
                                'Valor Unit. Supplier': "",
                                'Total Supplier': ""
                            })

            # --- LÓGICA SUPREME ---
            elif cliente == "Supreme":
                for aba in xl.sheet_names:
                    if "TOTAL" in aba.upper(): continue
                    df = xl.parse(aba, header=None)
                    tams = {c: str(df.iloc[14, c]) for c in range(9, 16) if pd.notna(df.iloc[14, c])}
                    
                    for start in range(16, len(df), 14):
                        dest = str(df.iloc[start, 0]).strip()
                        if not dest or dest == "nan": dest = "Geral"
                        
                        for i in range(start + 1, start + 13):
                            if i >= len(df) or pd.isna(df.iloc[i, 6]): continue
                            p = pd.to_numeric(df.iloc[i, 17], errors='coerce')
                            for c_idx, t_nom in tams.items():
                                q = pd.to_numeric(df.iloc[i, c_idx], errors='coerce')
                                if q and q > 0:
                                    lista_dados.append({
                                        'Referência': ref_manual, 
                                        'Designação': des_manual, 
                                        'Quant.': q, 
                                        'Pr.Unit.': 0, 
                                        'Pr.Unit.Moeda': p if pd.notna(p) else 0, 
                                        'Tabela de IVA': 4, 
                                        'Cor': df.iloc[i, 6], 
                                        'Tamanho': t_nom, 
                                        'TOTAL': q * (p if pd.notna(p) else 0), 
                                        'Destino': dest, 
                                        'Nr. CPO': "",
                                        'Nr. SPO': "",
                                        'Valor Unit. Supplier': "",
                                        'Total Supplier': ""
                                    })

        # --- LÓGICA STUDIO NICHOLSON ---
        elif arquivo.name.endswith('.pdf') and cliente == "Studio Nicholson":
            with pdfplumber.open(arquivo) as pdf:
                tams_ref = ["XXS", "XS", "S", "M", "L", "XL", "XXL", "UK4", "UK6", "UK8", "UK10", "UK12", "UK14"]
                lixo_geral = ["JERSEY", "MICRO", "RIB", "SHORT", "SLEEVE", "NECK", "VEST", "HENLEY", "COTTON", "BRANDED", "BOXY", "FIT", "T-SHIRT", "QTY", "COST", "TOTAL", "FIRST", "MAKE"]
                modelo_nicholson = ""
                destino_nicholson = "Ver PDF"
                for page in pdf.pages:
                    palavras = page.extract_words()
                    texto = page.extract_text()
                    if not texto: continue
                    linhas = texto.split('\n')
                    ship_match = re.search(r"Ship To:\s*(.*)", texto, re.IGNORECASE)
                    if ship_match:
                        linhas_ship = texto.split("Ship To:")[1].split('\n')
                        destino_nicholson = linhas_ship[1].strip() if len(linhas_ship) > 1 else destino_nicholson
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
                            modelo_nicholson = re.split(r"Qty|Cost|Total|First", linha, flags=re.I)[0].strip()
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
                                                lista_dados.append({
                                                    'Referência': "", 'Designação': modelo_nicholson, 'Quant.': q_num, 
                                                    'Pr.Unit.': p_unit, 'Pr.Unit.Moeda': 0, 'Tabela de IVA': 4, 
                                                    'Cor': cor_final, 'Tamanho': m['tam'], 'TOTAL': q_num * p_unit, 
                                                    'Destino': destino_nicholson, 'Nr. CPO': "", 
                                                    'Nr. SPO': "", 'Valor Unit. Supplier': "", 'Total Supplier': ""
                                                })

        # --- GERAÇÃO FINAL ---
        if lista_dados:
            df_final = pd.DataFrame(lista_dados).drop_duplicates()
            # Colunas com o novo nome "Nr. CPO"
            cols = ['Referência', 'Designação', 'Quant.', 'Pr.Unit.', 'Pr.Unit.Moeda', 'Tabela de IVA', 'Cor', 'Tamanho', 'TOTAL', 'Destino', 'Nr. CPO', 'Nr. SPO', 'Valor Unit. Supplier', 'Total Supplier']
            
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='openpyxl') as writer:
                for destino in df_final['Destino'].unique():
                    nome_aba = str(destino).replace('[','').replace(']','').replace('*','').replace(':','').replace('?','').replace('/','').replace('\\','')[:31]
                    df_dest = df_final[df_final['Destino'] == destino]
                    df_dest[cols].to_excel(writer, index=False, sheet_name=nome_aba)
            
            st.success(f"✅ Conversão concluída! Coluna 'Nr. CPO' atualizada.")
            st.download_button("⬇️ Descarregar Excel PHC", out.getvalue(), f"IMPORTAR_{cliente}.xlsx")
        else:
            st.warning("Nenhum dado encontrado.")

    except Exception as e:
        st.error(f"Erro no processamento: {e}")
