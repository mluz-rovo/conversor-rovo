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

        # --- LÓGICA PDF STUDIO NICHOLSON (DETETIVE DE LINHA) ---
        elif arquivo.name.endswith('.pdf') and cliente == "Studio Nicholson":
            with pdfplumber.open(arquivo) as pdf:
                tams_nich = ["UK4", "UK6", "UK8", "UK10", "UK12", "UK14", "XS", "S", "M", "L", "XL", "XXL"]
                
                for page in pdf.pages:
                    linhas = page.extract_text().split('\n')
                    destino = "Ver PDF"
                    modelo_atual = ""
                    
                    for idx, linha in enumerate(linhas):
                        linha_up = linha.upper()
                        
                        # 1. Destino
                        if "SHIP TO:" in linha_up:
                            destino = linhas[idx+1].strip() if idx+1 < len(linhas) else "Ver PDF"
                        
                        # 2. Modelo (Designação) - Ex: SORIN SNW-1868
                        if "SNW-" in linha_up or "SNM-" in linha_up:
                            modelo_atual = linha.strip()
                        
                        # 3. Deteção da Cor e Quantidades (Linha com €)
                        if "€" in linha:
                            # Preço
                            p_match = re.search(r"€\s*([\d,.]+)", linha)
                            p_moeda = pd.to_numeric(p_match.group(1).replace(',', ''), errors='coerce') if p_match else 0
                            
                            # Cor: Olhamos para a linha imediatamente acima
                            # No seu PDF, a cor (BLACK/PANNA) aparece sozinha antes dos números
                            cor_candidata = linhas[idx-1].strip() if idx > 0 else "Ver PDF"
                            
                            # Se a linha de cima for "Oty" ou "Qty", tentamos a anterior a essa
                            if cor_candidata.upper() in ["OTY", "QTY", "QUANTITY"]:
                                cor_candidata = linhas[idx-2].strip() if idx > 1 else "Ver PDF"

                            # Quantidades: Apanhamos os números na linha do €
                            partes = linha.split()
                            numeros = [n for n in partes if n.replace('.', '').isdigit()]
                            # Removemos o total da linha (penúltimo ou último antes do €)
                            qts_puras = numeros[:-1] if len(numeros) > 1 else numeros
                            
                            for i_q, val_q in enumerate(qts_puras):
                                if i_q < len(tams_nich):
                                    lista_dados.append({
                                        'Referência': "", 
                                        'Designação': modelo_atual, 
                                        'Quant.': pd.to_numeric(val_q, errors='coerce'), 
                                        'Pr.Unit.': 0, 
                                        'Pr.Unit.Moeda': p_moeda, 
                                        'Tabela de IVA': 4, 
                                        'Cor': cor_candidata, 
                                        'Tamanho': tams_nich[i_q], 
                                        'TOTAL': pd.to_numeric(val_q, errors='coerce') * p_moeda, 
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
            st.success(f"✅ Conversão concluída!")
            st.download_button("⬇️ Descarregar Excel", out.getvalue(), f"IMPORTAR_{cliente}.xlsx")
        else:
            st.warning("Dados não encontrados. Verifique se o PDF segue o padrão Nicholson.")

    except Exception as e:
        st.error(f"Erro: {e}")
