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

        # --- LÓGICA STUDIO NICHOLSON (MÉTODO DE TABELA RÍGIDA) ---
        elif arquivo.name.endswith('.pdf') and cliente == "Studio Nicholson":
            with pdfplumber.open(arquivo) as pdf:
                for page in pdf.pages:
                    texto = page.extract_text()
                    linhas = texto.split('\n')
                    
                    # 1. Capturar Destino e Modelo
                    destino = "Ver PDF"
                    modelo = ""
                    for i, l in enumerate(linhas):
                        if "Ship To:" in l and i+1 < len(linhas):
                            destino = linhas[i+1].strip()
                        if any(x in l.upper() for x in ["SNW -", "SNM -", "SN -", "LAY "]):
                            modelo = re.split(r"Qty|Cost|Total|First", l, flags=re.I)[0].strip()

                    # 2. Extrair tabelas de forma estruturada (Evita duplicar por coordenadas)
                    table = page.extract_table({
                        "vertical_strategy": "text", 
                        "horizontal_strategy": "text",
                        "snap_tolerance": 3,
                    })
                    
                    if not table: continue
                    
                    # Identificar cabeçalho de tamanhos
                    header = []
                    for row in table:
                        row_clean = [str(c).upper().strip() for c in row if c]
                        if any(t in row_clean for t in ["XS", "S", "M", "L", "XL", "UK4", "UK6"]):
                            header = [str(c).replace('\n', ' ').strip() for c in row]
                            break
                    
                    if not header: continue

                    for row in table:
                        row_str = " ".join([str(c) for c in row if c]).upper()
                        # Ignorar linhas de Totais e First Make
                        if "TOTAL" in row_str or "FIRST" in row_str or "€" not in row_str:
                            continue
                        
                        # Capturar Cor e Preço
                        cor = "Ver PDF"
                        preco = 0
                        for cell in row:
                            c_txt = str(cell).strip()
                            if "€" in c_txt:
                                p_match = re.findall(r"(\d+[\.,]\d{2})", c_txt)
                                preco = float(p_match[0].replace(',', '')) if p_match else 0
                            elif len(c_txt) > 2 and c_txt.isalpha() and c_txt not in ["QTY", "COST"]:
                                cor = c_txt

                        # Capturar Quantidades mapeadas pelo cabeçalho
                        for idx, cell in enumerate(row):
                            if idx < len(header):
                                tam_nome = header[idx].upper()
                                if any(t == tam_nome or (t in tam_nome and "/" in tam_nome) for t in ["XXS", "XS", "S", "M", "L", "XL", "XXL", "UK4", "UK6", "UK8", "UK10", "UK12", "UK14"]):
                                    val_q = str(cell).strip()
                                    if val_q.isdigit():
                                        q_num = int(val_q)
                                        if q_num > 0:
                                            lista_dados.append({
                                                'Referência': "", 'Designação': modelo, 'Quant.': q_num,
                                                'Pr.Unit.': preco, 'Pr.Unit.Moeda': 0, 'Tabela de IVA': 4,
                                                'Cor': cor, 'Tamanho': header[idx], 'TOTAL': q_num * preco, 
                                                'Destino': destino, 'CPO': ""
                                            })

        if lista_dados:
            df = pd.DataFrame(lista_dados)
            # Remove duplicados exatos para garantir
            df = df.drop_duplicates(subset=['Designação', 'Quant.', 'Cor', 'Tamanho', 'TOTAL'])
            cols = ['Referência', 'Designação', 'Quant.', 'Pr.Unit.', 'Pr.Unit.Moeda', 'Tabela de IVA', 'Cor', 'Tamanho', 'TOTAL', 'Destino', 'CPO']
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='openpyxl') as writer:
                df[cols].to_excel(writer, index=False, sheet_name="PHC")
            st.success("✅ Conversão Studio Nicholson afinada!")
            st.download_button("⬇️ Descarregar Excel", out.getvalue(), "IMPORTAR_PHC.xlsx")
    except Exception as e:
        st.error(f"Erro: {e}")
