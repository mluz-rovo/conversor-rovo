import streamlit as st
import pandas as pd
import io
import pdfplumber
import re

# Versão do Código: 2.1 - Correção Erro Index Coluna Stussy
st.set_page_config(page_title="ROVO - Conversor Universal", page_icon="🚀")
st.title("🚀 ROVO Universal Converter")

cliente = st.sidebar.selectbox("Selecione o Cliente", ["Stussy", "Supreme", "Studio Nicholson"])
st.sidebar.write("---")
st.sidebar.info("💡 Se as alterações não aparecerem, faça Refresh (F5) na página.")

st.info(f"Modo Ativo: **{cliente}**")

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
                    # Proteção: Verifica se a linha tem colunas suficientes (Erro 17 evitado aqui)
                    if len(row) > 17:
                        q = pd.to_numeric(row[12], errors='coerce')
                        p = pd.to_numeric(row[17], errors='coerce')
                        if q and q > 0:
                            lista_dados.append({
                                'Referência': "", 
                                'Designação': "", 
                                'Quant.': q, 
                                'Pr.Unit.': 0, 
                                'Pr.Unit.Moeda': p if pd.notna(p) else 0, 
                                'Tabela de IVA': 4, 
                                'Cor': row[6] if len(row) > 6 else "", 
                                'Tamanho': row[9] if len(row) > 9 else "", 
                                'TOTAL': q * (p if pd.notna(p) else 0), 
                                'Destino': row[4] if len(row) > 4 else "", 
                                'Aba': "Stussy_PO"
                            })
                    else:
                        continue # Salta linhas que não têm o formato esperado

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

        # --- LÓGICA PDF STUDIO NICHOLSON ---
        elif arquivo.name.endswith('.pdf') and cliente == "Studio Nicholson":
            with pdfplumber.open(arquivo) as pdf:
                tams_ref = ["XS", "S", "M", "L", "XL", "XXL", "UK4", "UK6", "UK8", "UK10", "UK12", "UK14"]
                palavras_proibidas = ["JERSEY", "MICRO", "RIB", "MERCERIZED", "COTTON", "BRANDED", "BOXY", "FIT", "T-SHIRT", "VEST", "HENLEY", "SCOOP", "NECK", "TOTAL", "QTY", "OTY", "SNM", "SNW", "LAY", "PRODUCTION", "ORDER"]

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
                        for r_idx, row in enumerate(table):
                            row_str = " ".join([str(x).upper() for x in row if x])
                            if any(t in row_str for t in tams_ref):
                                headers = [str(x).replace('\n', ' ').strip() for x in row]
                                start_data = r_idx + 1
                                break
                        
                        if start_data == -1: continue

                        for i in range(start_data, len(table)):
                            row_data = table[i]
                            row_str_full = " ".join([str(x) for x in row_data if x]).replace('\n', ' ')
                            
                            if "€" in row_str_full:
                                designacao = str(row_data[0]).split('\n')[0].strip()
                                partes = row_str_full.split()
                                cor_limpa = [p for p in partes if p.upper().replace(',', '').replace('.', '') not in palavras_proibidas and not p.replace('.', '').isdigit() and "€" not in p and "SN" not in p.upper() and len(p) > 2]
                                cor_resultado = " ".join(cor_limpa).strip()
                                
                                p_valor = 0
                                for cell in row_data:
                                    if "€" in str(cell):
                                        p_txt = str(cell).replace('€','').replace(',','.').replace(' ', '').strip()
                                        p_valor = pd.to_numeric(p_txt, errors='coerce')
                                        break

                                for col_idx, h_text in enumerate(headers):
                                    tam_detectado = next((t for t in tams_ref if t in h_text.upper()), "")
                                    if tam_detectado:
                                        qtd = pd.to_numeric(row_data[col_idx], errors='coerce')
                                        if qtd and qtd > 0:
                                            lista_dados.append({
                                                'Referência': "", 'Designação': designacao, 'Quant.': qtd, 'Pr.Unit.': p_valor, 'Pr.Unit.Moeda': 0, 'Tabela de IVA': 4, 'Cor': cor_resultado, 'Tamanho': h_text, 'TOTAL': qtd * p_valor, 'Destino': destino, 'Aba': "Nicholson_PO"
                                            })

        if lista_dados:
            df_final = pd.DataFrame(lista_dados)
            df_final['CPO'] = ""
            cols = ['Referência', 'Designação', 'Quant.', 'Pr.Unit.', 'Pr.Unit.Moeda', 'Tabela de IVA', 'Cor', 'Tamanho', 'TOTAL', 'Destino', 'CPO']
            
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='openpyxl') as writer:
                for aba_nom in df_final['Aba'].unique():
                    df_final[df_final['Aba'] == aba_nom][cols].to_excel(writer, sheet_name=str(aba_nom)[:31], index=False)
            
            st.success(f"✅ Conversão concluída para {cliente}!")
            st.download_button("⬇️ Descarregar Excel PHC", out.getvalue(), f"IMPORTAR_{cliente}.xlsx")
        else:
            st.warning("Nenhum dado válido encontrado no ficheiro.")

    except Exception as e:
        st.error(f"Erro no processamento: {e}")
