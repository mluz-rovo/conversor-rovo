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

        # --- LÓGICA PDF (STUDIO NICHOLSON) - AJUSTADA AO TEU PDF ---
        elif arquivo.name.endswith('.pdf') and cliente == "Studio Nicholson":
            with pdfplumber.open(arquivo) as pdf:
                for page in pdf.pages:
                    linhas = page.extract_text().split('\n')
                    
                    destino = "Ver PDF"
                    modelo_atual = ""
                    nomes_tams = ["UK4", "UK6", "UK8", "UK10", "UK12", "UK14"]
                    
                    for idx, linha in enumerate(linhas):
                        # 1. Capturar Destino (Ship To)
                        if "Ship To:" in linha:
                            destino = linhas[idx+1].strip() if idx+1 < len(linhas) else "Ver PDF"
                        
                        # 2. Capturar Modelo (Ex: SORIN SNW-1868)
                        if "SNW-" in linha or "SNM-" in linha:
                            modelo_atual = linha.split(' ')[0] + " " + linha.split(' ')[1]
                        
                        # 3. Capturar Linha com Quantidades e Preço
                        if "€" in linha and any(char.isdigit() for char in linha):
                            partes = linha.split()
                            
                            # A cor costuma ser a primeira palavra da linha (ex: BLACK, PANNA)
                            cor = partes[0]
                            
                            # O preço é o valor que vem a seguir ao símbolo € (ex: € 29.45)
                            p_moeda = 0
                            if "€" in partes:
                                p_idx = partes.index("€")
                                p_moeda = pd.to_numeric(partes[p_idx+1].replace(',',''), errors='coerce')
                            
                            # Capturar as quantidades (são os números antes da Qty/Total)
                            # No seu PDF: "11 9 8 6 3 37 € 29.45"
                            qts_candidatas = [pd.to_numeric(p, errors='coerce') for p in partes if p.isdigit()]
                            # Removemos o último número que é o total da linha (ex: 37)
                            qts_puras = qts_candidatas[:-1] if len(qts_candidatas) > 1 else qts_candidatas
                            
                            for i_q, valor_q in enumerate(qts_puras):
                                if i_q < len(nomes_tams):
                                    lista_dados.append({
                                        'Referência': modelo_atual, 
                                        'Designação': "", 
                                        'Quant.': valor_q, 
                                        'Pr.Unit.': 0, 
                                        'Pr.Unit.Moeda': p_moeda, 
                                        'Tabela de IVA': 4, 
                                        'Cor': cor, 
                                        'Tamanho': nomes_tams[i_q], 
                                        'TOTAL': valor_q * (p_moeda if p_moeda else 0), 
                                        'Destino': destino, 
                                        'Aba': "Nicholson_PO"
                                    })

        # --- EXPORTAÇÃO ---
        df_final = pd.DataFrame(lista_dados)
        if not df_final.empty:
            df_final['CPO'] = ""
            cols = ['Referência', 'Designação', 'Quant.', 'Pr.Unit.', 'Pr.Unit.Moeda', 'Tabela de IVA', 'Cor', 'Tamanho', 'TOTAL', 'Destino', 'CPO']
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='openpyxl') as writer:
                for aba_nom in df_final['Aba'].unique():
                    df_final[df_final['Aba'] == aba_nom][cols].to_excel(writer, sheet_name=str(aba_nom)[:31], index=False)
            st.success(f"✅ Sucesso! Convertidas {len(df_final)} linhas.")
            st.download_button("⬇️ Descarregar Excel", out.getvalue(), f"IMPORTAR_{cliente}.xlsx")
        else:
            st.warning("Dados não encontrados. Verifique se o PDF segue o padrão Nicholson.")

    except Exception as e:
        st.error(f"Erro: {e}")
