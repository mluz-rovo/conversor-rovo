import streamlit as st
import pandas as pd
import io
import pdfplumber
import re

# Configuração da página (Voltamos ao visual padrão, sem o verde)
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
            # (Mantemos a lógica que já funcionava para Stussy e Supreme)
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

        # --- LÓGICA PDF STUDIO NICHOLSON (FINAL E CORRIGIDA) ---
        elif arquivo.name.endswith('.pdf') and cliente == "Studio Nicholson":
            with pdfplumber.open(arquivo) as pdf:
                tams_ref = ["XS", "S", "M", "L", "XL", "XXL", "UK4", "UK6", "UK8", "UK10", "UK12", "UK14"]
                proibidos = ["JERSEY", "MICRO", "RIB", "MERCERIZED", "COTTON", "BRANDED", "BOXY", "FIT", "T-SHIRT", "SNW", "SNM", "LAY", "QTY", "OTY", "TOTAL"]

                for page in pdf.pages:
                    texto_pg = page.extract_text() or ""
                    tabelas = page.extract_tables()
                    
                    ship = re.search(r"Ship To:\s*(.*)", texto_pg, re.IGNORECASE)
                    destino = ship.group(1).split('\n')[0].strip() if ship else "Ver PDF"

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
                                # Designação recebe o Modelo (Coluna 1 da tabela)
                                modelo = str(row_data[0]).split('\n')[0].strip()
                                
                                # Limpeza de Cor (Ignora JERSEY e derivados)
                                partes = row_str_full.split()
                                cor_limpa = [p for p in partes if p.upper() not in proibidos and not p.replace('.','').isdigit() and "€" not in p and "SN" not in p.upper() and len(p)>2]
                                cor_final = " ".join(cor_limpa).strip()
                                
                                # Preço para Pr.Unit.
                                p_unit = 0
                                for cell in row_data:
                                    if "€" in str(cell):
                                        p_txt = str(cell).replace('€','').replace(',','.').replace(' ', '').strip()
                                        p_unit = pd.to_numeric(p_txt, errors='coerce')
                                        break

                                for col_idx, h_text in enumerate(headers):
                                    tam_ok = ""
                                    for t in tams_ref:
                                        if t in h_text.upper():
                                            tam_ok = h_text
                                            break
                                    
                                    if tam_ok:
                                        qtd = pd.to_numeric(row_data[col_idx], errors='coerce')
                                        if qtd and qtd > 0:
                                            lista_dados.append({
                                                'Referência': "", 
                                                'Designação': modelo, 
                                                'Quant.': qtd, 
                                                'Pr.Unit.': p_unit, 
                                                'Pr.Unit.Moeda': 0, 
                                                'Tabela de IVA': 4, 
                                                'Cor': cor_final, 
                                                'Tamanho': tam_ok, 
                                                'TOTAL': qtd * p_unit, 
                                                'Destino': destino, 
                                                'CPO': ""
                                            })

        # --- EXPORTAÇÃO ---
        if lista_dados:
            df_final = pd.DataFrame(lista_dados)
            cols = ['Referência', 'Designação', 'Quant.', 'Pr.Unit.', 'Pr.Unit.Moeda', 'Tabela de IVA', 'Cor', 'Tamanho', 'TOTAL', 'Destino', 'CPO']
            
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='openpyxl') as writer:
                df_final[cols].to_excel(writer, index=False, sheet_name="Importar PHC")
            
            st.success(f"✅ Conversão de {cliente} concluída!")
            st.download_button("⬇️ Descarregar Excel para PHC", out.getvalue(), f"IMPORTAR_{cliente}.xlsx")
        else:
            st.warning("Nenhum dado encontrado no ficheiro.")

    except Exception as e:
        st.error(f"Erro no processamento: {e}")
