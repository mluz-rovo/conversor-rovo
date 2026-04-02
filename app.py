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

        # --- CASO 1: EXCEL (Stussy ou Supreme) ---
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

        # --- CASO 2: PDF (Studio Nicholson) com pdfplumber (SEM JAVA) ---
        elif arquivo.name.endswith('.pdf') and cliente == "Studio Nicholson":
            with pdfplumber.open(arquivo) as pdf:
                for page in pdf.pages:
                    texto = page.extract_text()
                    tabelas = page.extract_tables()
                    
                    # Tentar apanhar o Destino (Ship To) no texto da página
                    destino = "Ver PDF"
                    ship_match = re.search(r"Ship To:\s*(.*)", texto)
                    if ship_match:
                        destino = ship_match.group(1).split('\n')[0].strip()

                    for table in tabelas:
                        modelo = ""
                        for row in table:
                            # row é uma lista de colunas
                            row_str = " ".join([str(x) for x in row if x])
                            
                            # Detetar Modelo
                            if "SNW -" in row_str or "SNM -" in row_str:
                                modelo = str(row[0]).split('\n')[0].strip()
                            
                            # Detetar Linha de dados (tem o preço com €)
                            if "€" in row_str:
                                # Ajustar índices baseado no comportamento do pdfplumber
                                cor = str(row[1]).strip() if row[1] else ""
                                p_texto = row_str.split('€')[-2].split()[-1].replace(',','.')
                                p_moeda = pd.to_numeric(p_texto, errors='coerce')
                                
                                # Tamanhos (UK4 a UK14 costumam aparecer nas colunas centrais)
                                nomes_tams = ["UK4", "UK6", "UK8", "UK10", "UK12", "UK14"]
                                # Tentamos mapear as colunas onde aparecem números
                                col_tams = [i for i, x in enumerate(row) if str(x).isdigit()]
                                
                                for i_col, idx_real in enumerate(col_tams):
                                    if i_col < len(nomes_tams):
                                        q = pd.to_numeric(row[idx_real], errors='coerce')
                                        if q and q > 0:
                                            lista_dados.append({
                                                'Referência': modelo, 'Designação': "", 'Quant.': q, 'Pr.Unit.': 0, 
                                                'Pr.Unit.Moeda': p_moeda, 'Tabela de IVA': 4, 'Cor': cor, 
                                                'Tamanho': nomes_tams[i_col], 'TOTAL': q * (p_moeda if p_moeda else 0), 
                                                'Destino': destino, 'Aba': "Nicholson_PO"
                                            })

        # --- FINALIZAÇÃO ---
        df_final = pd.DataFrame(lista_dados)
        if not df_final.empty:
            df_final['CPO'] = ""
            cols = ['Referência', 'Designação', 'Quant.', 'Pr.Unit.', 'Pr.Unit.Moeda', 'Tabela de IVA', 'Cor', 'Tamanho', 'TOTAL', 'Destino', 'CPO']
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='openpyxl') as writer:
                for aba_nom in df_final['Aba'].unique():
                    df_final[df_final['Aba'] == aba_nom][cols].to_excel(writer, sheet_name=str(aba_nom)[:31], index=False)
            st.success("✅ Convertido!")
            st.download_button("⬇️ Descarregar Excel", out.getvalue(), f"IMPORTAR_{cliente}.xlsx")
        else:
            st.warning("Dados não encontrados. Verifique o cliente selecionado.")

    except Exception as e:
        st.error(f"Erro: {e}")
