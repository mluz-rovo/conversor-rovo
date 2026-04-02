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

        # --- LÓGICA PDF STUDIO NICHOLSON (HÍBRIDA UK/IT + XS/XXL) ---
        elif arquivo.name.endswith('.pdf') and cliente == "Studio Nicholson":
            with pdfplumber.open(arquivo) as pdf:
                # Super Lista de Tamanhos Possíveis
                tams_possiveis = ["XS", "S", "M", "L", "XL", "XXL", "UK4", "UK6", "UK8", "UK10", "UK12", "UK14", "IT36", "IT38", "IT40", "IT42", "IT44", "IT46"]
                
                for page in pdf.pages:
                    texto_pg = page.extract_text() or ""
                    tabelas = page.extract_tables()
                    
                    # 1. Capturar Destino (Ship To)
                    destino = "Ver PDF"
                    ship_match = re.search(r"Ship To:\s*(.*)", texto_pg, re.IGNORECASE)
                    if ship_match:
                        destino = ship_match.group(1).split('\n')[0].strip()

                    for table in tabelas:
                        headers = []
                        start_data = -1
                        
                        # 2. Encontrar Linha de Cabeçalho (identifica onde estão os tamanhos)
                        for row_idx, row in enumerate(table):
                            row_str = " ".join([str(x).upper() for x in row if x])
                            # Verifica se a linha contém pelo menos um dos tamanhos da nossa super lista
                            if any(t in row_str for t in tams_possiveis):
                                headers = [str(x).replace('\n', ' ').strip() for x in row]
                                start_data = row_idx + 1
                                break
                        
                        if start_data == -1: continue

                        # 3. Processar Linhas de Dados
                        for i in range(start_data, len(table)):
                            row_data = table[i]
                            row_str = " ".join([str(x) for x in row_data if x])
                            
                            if "€" in row_str:
                                # Tenta capturar Designação (primeira coluna) e Cor (segunda coluna)
                                designacao = str(row_data[0]).split('\n')[0] if row_data[0] else "Nicholson Item"
                                cor = str(row_data[1]).split('\n')[-1] if len(row_data) > 1 else "Ver PDF"
                                
                                # Extrair Preço Unitário
                                p_moeda = 0
                                for cell in row_data:
                                    if "€" in str(cell):
                                        p_txt = str(cell).replace('€','').replace(',','.').replace(' ', '').strip()
                                        p_moeda = pd.to_numeric(p_txt, errors='coerce')
                                        break
                                
                                # 4. Mapear Quantidades com base nos headers detetados
                                for col_idx, h_text in enumerate(headers):
                                    # Verifica se este cabeçalho é um tamanho conhecido
                                    tamanho_limpo = ""
                                    for t in tams_possiveis:
                                        if t in h_text.upper():
                                            tamanho_limpo = t
                                            break
                                    
                                    if tamanho_limpo:
                                        qtd = pd.to_numeric(row_data[col_idx], errors='coerce')
                                        if qtd and qtd > 0:
                                            lista_dados.append({
                                                'Referência': "", 
                                                'Designação': designacao, 
                                                'Quant.': qtd, 
                                                'Pr.Unit.': 0, 
                                                'Pr.Unit.Moeda': p_moeda, 
                                                'Tabela de IVA': 4, 
                                                'Cor': cor, 
                                                'Tamanho': h_text, # Mantém o texto original do PDF (ex: UK12/IT44)
                                                'TOTAL': qtd * (p_moeda if p_moeda else 0), 
                                                'Destino': destino, 
                                                'Aba': "Nicholson_PO"
                                            })

        # --- EXPORTAÇÃO FINAL ---
        df_final = pd.DataFrame(lista_dados)
        if not df_final.empty:
            df_final['CPO'] = ""
            cols = ['Referência', 'Designação', 'Quant.', 'Pr.Unit.', 'Pr.Unit.Moeda', 'Tabela de IVA', 'Cor', 'Tamanho', 'TOTAL', 'Destino', 'CPO']
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='openpyxl') as writer:
                for aba_nom in df_final['Aba'].unique():
                    df_final[df_final['Aba'] == aba_nom][cols].to_excel(writer, sheet_name=str(aba_nom)[:31], index=False)
            st.success(f"✅ Conversão concluída! Processadas {len(df_final)} linhas.")
            st.download_button("⬇️ Descarregar Excel para PHC", out.getvalue(), f"IMPORTAR_{cliente}.xlsx")
        else:
            st.warning("Dados não encontrados. Verifique se o PDF tem a tabela visível ou se o cliente está correto.")

    except Exception as e:
        st.error(f"Erro técnico: {e}")
