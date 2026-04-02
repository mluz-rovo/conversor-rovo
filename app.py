import streamlit as st
import pandas as pd
import io
import pdfplumber
import re

st.set_page_config(page_title="ROVO - Conversor Universal", page_icon="🚀")
st.title("🚀 ROVO Universal Converter")

cliente = st.sidebar.selectbox("Selecione o Cliente", ["Stussy", "Supreme", "Studio Nicholson"])
st.info(f"Modo: **{cliente}**")

formatos = ["xlsx", "pdf"] if cliente == "Studio Nicholson" else ["xlsx"]
arquivo = st.file_uploader(f"Carregar ficheiro", type=formatos)

if arquivo:
    try:
        lista_dados = []

        # --- LÓGICA EXCEL (STUSSY / SUPREME) ---
        if arquivo.name.endswith('.xlsx'):
            xl = pd.ExcelFile(arquivo, engine='openpyxl')
            if cliente == "Stussy":
                df = xl.parse(xl.sheet_names[0], header=None)
                for i, row in df.iloc[1:].iterrows():
                    q, p = pd.to_numeric(row[12], errors='coerce'), pd.to_numeric(row[17], errors='coerce')
                    if q and q > 0:
                        lista_dados.append({'Referência': "", 'Designação': "", 'Quant.': q, 'Pr.Unit.': 0, 'Pr.Unit.Moeda': p, 'Tabela de IVA': 4, 'Cor': row[6], 'Tamanho': row[9], 'TOTAL': q*(p if p else 0), 'Destino': row[4], 'Aba': "Stussy_PO"})
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
                                    lista_dados.append({'Referência': "", 'Designação': "", 'Quant.': q, 'Pr.Unit.': 0, 'Pr.Unit.Moeda': p, 'Tabela de IVA': 4, 'Cor': df.iloc[i, 6], 'Tamanho': t_nom, 'TOTAL': q*(p if p else 0), 'Destino': dest, 'Aba': aba})

        # --- LÓGICA PDF STUDIO NICHOLSON (LIMPEZA RADICAL DE COR) ---
        elif arquivo.name.endswith('.pdf') and cliente == "Studio Nicholson":
            with pdfplumber.open(arquivo) as pdf:
                tams_possiveis = ["XS", "S", "M", "L", "XL", "XXL", "UK4", "UK6", "UK8", "UK10", "UK12", "UK14"]
                # Lista de palavras que NUNCA são cores
                lixo_textual = ["JERSEY", "MICRO", "RIB", "COTTON", "MERCERIZED", "BRANDED", "BOXY", "FIT", "T-SHIRT", "VEST", "HENLEY", "SHORT", "SLEEVE", "SCOOP", "NECK", "OPTIC", "Oty", "Qty", "FIRST/MAKE", "COST", "TOTAL", "SNW", "SNM", "LAY", "PRODUCTION", "ORDER"]

                for page in pdf.pages:
                    texto_pg = page.extract_text() or ""
                    tabelas = page.extract_tables()
                    
                    destino = "Ver PDF"
                    ship_match = re.search(r"Ship To:\s*(.*)", texto_pg, re.IGNORECASE)
                    if ship_match:
                        destino = ship_match.group(1).split('\n')[0].strip()

                    for table in tabelas:
                        headers = []
                        start_data = -1
                        for row_idx, row in enumerate(table):
                            row_str = " ".join([str(x).upper() for x in row if x])
                            if any(t in row_str for t in tams_possiveis):
                                headers = [str(x).replace('\n', ' ').strip() for x in row]
                                start_data = row_idx + 1
                                break
                        
                        if start_data == -1: continue

                        for i in range(start_data, len(table)):
                            row_data = table[i]
                            # row_str_completa junta todas as células daquela linha de dados
                            row_str_completa = " ".join([str(x) for x in row_data if x]).replace('\n', ' ')
                            
                            if "€" in row_str_completa:
                                # 1. DESIGNACAO: Pegamos na primeira célula e limpamos códigos
                                designacao = str(row_data[0]).split('\n')[0].strip()
                                
                                # 2. COR (A Nova Lógica Radical):
                                # Pegamos no texto bruto da linha
                                partes = row_str_completa.split()
                                cor_limpa = []
                                for p in partes:
                                    p_clean = p.upper().replace(',', '').replace('.', '')
                                    # Só aceitamos se: não for lixo, não for número, não tiver €, não for o modelo
                                    if (p_clean not in lixo_textual and 
                                        not p_clean.isdigit() and 
                                        "€" not in p_clean and 
                                        "SNW" not in p_clean and 
                                        "SNM" not in p_clean and
                                        len(p_clean) > 2):
                                        cor_limpa.append(p)
                                
                                # O resultado deve ser "BLACK" ou "PANNA" ou "OPTIC WHITE"
                                cor_final = " ".join(cor_limpa).strip()
                                
                                # 3. PREÇO
                                p_moeda = 0
                                for cell in row_data:
                                    if "€" in str(cell):
                                        p_txt = str(cell).replace('€','').replace(',','.').replace(' ', '').strip()
                                        p_moeda = pd.to_numeric(p_txt, errors='coerce')
                                        break
                                
                                # 4. QUANTIDADES
                                for col_idx, h_text in enumerate(headers):
                                    tamanho_encontrado = ""
                                    for t in tams_possiveis:
                                        if t in h_text.upper():
                                            tamanho_encontrado = h_text
                                            break
                                    
                                    if tamanho_encontrado:
                                        qtd = pd.to_numeric(row_data[col_idx], errors='coerce')
                                        if qtd and qtd > 0:
                                            lista_dados.append({
                                                'Referência': "", 'Designação': designacao, 'Quant.': qtd, 
                                                'Pr.Unit.': 0, 'Pr.Unit.Moeda': p_moeda, 'Tabela de IVA': 4, 
                                                'Cor': cor_final, 'Tamanho': tamanho_encontrado, 
                                                'TOTAL': qtd * (p_moeda if p_moeda else 0), 'Destino': destino, 'Aba': "Nicholson_PO"
                                            })

        df_final = pd.DataFrame(lista_dados)
        if not df_final.empty:
            df_final['CPO'] = ""
            cols = ['Referência', 'Designação', 'Quant.', 'Pr.Unit.', 'Pr.Unit.Moeda', 'Tabela de IVA', 'Cor', 'Tamanho', 'TOTAL', 'Destino', 'CPO']
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='openpyxl') as writer:
                for aba_nom in df_final['Aba'].unique():
                    df_final[df_final['Aba'] == aba_nom][cols].to_excel(writer, sheet_name=str(aba_nom)[:31], index=False)
            st.success("✅ Conversão Nicholson concluída com limpeza radical!")
            st.download_button("⬇️ Descarregar Excel PHC", out.getvalue(), f"IMPORTAR_{cliente}.xlsx")
        else:
            st.warning("Dados não encontrados.")
    except Exception as e:
        st.error(f"Erro: {e}")
