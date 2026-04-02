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

        # --- LÓGICA PDF (STUDIO NICHOLSON) ---
        elif arquivo.name.endswith('.pdf') and cliente == "Studio Nicholson":
            with pdfplumber.open(arquivo) as pdf:
                for page in pdf.pages:
                    linhas_texto = page.extract_text().split('\n')
                    
                    destino = "Ver PDF"
                    modelo_detetado = ""
                    nomes_tams = ["UK4", "UK6", "UK8", "UK10", "UK12", "UK14"]
                    
                    for idx, linha in enumerate(linhas_texto):
                        # 1. Capturar Destino (Ship To)
                        if "Ship To:" in linha:
                            destino = linhas_texto[idx+1].strip() if idx+1 < len(linhas_texto) else "Ver PDF"
                        
                        # 2. Capturar Modelo (Vai para DESIGNAÇÃO)
                        if "SNW-" in linha or "SNM-" in linha:
                            # Pega as primeiras palavras que formam o modelo (ex: SORIN SNW-1868)
                            palavras_modelo = linha.split()
                            modelo_detetado = " ".join(palavras_modelo[:3])
                            continue
                        
                        # 3. Capturar Linha com Cor e Quantidades
                        if "€" in linha:
                            partes = linha.split()
                            
                            # A cor é a primeira palavra (BLACK, PANNA, etc)
                            cor_extraida = partes[0]
                            
                            # Preço unitário (€)
                            p_moeda = 0
                            if "€" in partes:
                                try:
                                    p_idx = partes.index("€")
                                    p_moeda = pd.to_numeric(partes[p_idx+1].replace(',',''), errors='coerce')
                                except: pass
                            
                            # Quantidades: filtramos apenas números antes do total da linha
                            numeros = [p for p in partes if p.replace('.','').isdigit()]
                            # A Nicholson mete o total da linha antes do €, então removemos o último número (total qty)
                            qts_reais = numeros[:-1] if len(numeros) > 1 else numeros
                            
                            for i_q, val_q in enumerate(qts_reais):
                                q_num = pd.to_numeric(val_q, errors='coerce')
                                if q_num and q_num > 0 and i_q < len(nomes_tams):
                                    lista_dados.append({
                                        'Referência': "", # Podes colocar algo aqui se quiseres
                                        'Designação': modelo_detetado, 
                                        'Quant.': q_num, 
                                        'Pr.Unit.': 0, 
                                        'Pr.Unit.Moeda': p_moeda, 
                                        'Tabela de IVA': 4, 
                                        'Cor': cor_extraida, 
                                        'Tamanho': nomes_tams[i_q], 
                                        'TOTAL': q_num * (p_moeda if p_moeda else 0), 
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
            st.success(f"✅ Conversão de {cliente} concluída!")
            st.download_button("⬇️ Descarregar Excel PHC", out.getvalue(), f"IMPORTAR_{cliente}.xlsx")
        else:
            st.warning("Ainda não conseguimos ler os dados. Verifique se o PDF é o original da Studio Nicholson.")

    except Exception as e:
        st.error(f"Erro: {e}")
