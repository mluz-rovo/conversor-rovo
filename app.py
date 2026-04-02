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

        # --- LÓGICA PDF STUDIO NICHOLSON (PREÇO NO PR.UNIT + COR CORRIGIDA) ---
        elif arquivo.name.endswith('.pdf') and cliente == "Studio Nicholson":
            with pdfplumber.open(arquivo) as pdf:
                tams_nich = ["UK4/IT36", "UK6/IT38", "UK8/IT40", "UK10/IT42", "UK12/IT44", "UK14/IT46", "XS", "S", "M", "L", "XL", "XXL"]
                
                for page in pdf.pages:
                    linhas = page.extract_text().split('\n')
                    destino = "Ver PDF"
                    modelo_atual = ""
                    
                    for idx, linha in enumerate(linhas):
                        linha_up = linha.upper()
                        
                        # 1. Destino (Ship To)
                        if "SHIP TO:" in linha_up:
                            destino = linhas[idx+1].strip() if idx+1 < len(linhas) else "Ver PDF"
                        
                        # 2. Capturar Modelo para DESIGNAÇÃO
                        if "SNW-" in linha_up or "SNM-" in linha_up:
                            modelo_atual = linha.strip()
                            continue
                        
                        # 3. Deteção da Linha de Dados (€)
                        if "€" in linha:
                            partes = linha.split()
                            
                            # Preço vai para PR.UNIT (conforme pedido)
                            p_match = re.search(r"€\s*([\d,.]+)", linha)
                            p_valor = pd.to_numeric(p_match.group(1).replace(',', ''), errors='coerce') if p_match else 0
                            
                            # --- NOVA LÓGICA DE COR ---
                            # Se a linha começa com a cor (ex: BLACK 11 9 8...), partes[0] é a cor.
                            cor_final = partes[0]
                            # Se a cor for "JERSEY" ou "MICRO", tentamos a palavra seguinte
                            if cor_final in ["JERSEY", "MICRO", "RIB", "QTY", "OTY"]:
                                cor_final = partes[1] if len(partes) > 1 else cor_final
                            
                            # Quantidades
                            numeros = [n for n in partes if n.replace('.', '').isdigit()]
                            qts_reais = numeros[:-1] if len(numeros) > 1 else numeros
                            
                            for i_q, val_q in enumerate(qts_reais):
                                if i_q < len(tams_nich):
                                    q_num = pd.to_numeric(val_q, errors='coerce')
                                    lista_dados.append({
                                        'Referência': "", 
                                        'Designação': modelo_atual, 
                                        'Quant.': q_num, 
                                        'Pr.Unit.': p_valor, # Preço aqui
                                        'Pr.Unit.Moeda': 0,   # Moeda a zero
                                        'Tabela de IVA': 4, 
                                        'Cor': cor_final, 
                                        'Tamanho': tams_nich[i_q], 
                                        'TOTAL': q_num * p_valor, 
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
            st.success("✅ Conversão Nicholson corrigida!")
            st.download_button("⬇️ Descarregar Excel PHC", out.getvalue(), f"IMPORTAR_{cliente}.xlsx")
        else:
            st.warning("Dados não encontrados.")

    except Exception as e:
        st.error(f"Erro: {e}")
