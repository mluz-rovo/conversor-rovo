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
                    texto_completo = page.extract_text()
                    tabelas = page.extract_tables()
                    
                    # Tentar apanhar o Destino
                    destino = "Ver PDF"
                    ship_match = re.search(r"Ship To:\s*(.*)", texto_completo, re.IGNORECASE)
                    if ship_match:
                        destino = ship_match.group(1).split('\n')[0].strip()

                    for table in tabelas:
                        modelo_atual = ""
                        for row in table:
                            # Limpar a linha de valores nulos
                            row_clean = [str(x).strip() if x else "" for x in row]
                            row_str = " ".join(row_clean)
                            
                            # 1. Detetar Modelo
                            if "SNW -" in row_str or "SNM -" in row_str:
                                modelo_atual = row_clean[0].split('\n')[0]
                                continue
                            
                            # 2. Detetar Linha de Produção (Preço com €)
                            if "€" in row_str:
                                # Preço: Procuramos o valor que tem o € colado ou perto
                                p_moeda = 0
                                for celula in row_clean:
                                    if "€" in celula:
                                        p_txt = celula.replace('€','').replace(',','.').strip()
                                        p_moeda = pd.to_numeric(p_txt, errors='coerce')
                                        break
                                
                                # Cor: Costuma ser o primeiro texto longo após o modelo
                                cor = row_clean[1] if len(row_clean) > 1 else ""
                                
                                # Tamanhos: UK4(Col 10), UK6(11), UK8(12), UK10(13), UK12(14), UK14(15)
                                nomes_tams = ["UK4", "UK6", "UK8", "UK10", "UK12", "UK14"]
                                # No pdfplumber as colunas de tamanhos costumam começar no índice 10
                                start_col = 10 
                                
                                for i_tam, nome_tam in enumerate(nomes_tams):
                                    col_idx = start_col + i_tam
                                    if col_idx < len(row_clean):
                                        qtd = pd.to_numeric(row_clean[col_idx], errors='coerce')
                                        if qtd and qtd > 0:
                                            lista_dados.append({
                                                'Referência': modelo_atual, 'Designação': "", 'Quant.': qtd, 
                                                'Pr.Unit.': 0, 'Pr.Unit.Moeda': p_moeda, 'Tabela de IVA': 4, 
                                                'Cor': cor, 'Tamanho': nome_tam, 
                                                'TOTAL': qtd * (p_moeda if p_moeda else 0), 
                                                'Destino': destino, 'Aba': "Nicholson_PO"
                                            })

        # --- EXPORTAÇÃO ---
        df_final = pd.DataFrame(lista_dados)
        if not df_final.empty:
            df_final['CPO'] = ""
            cols = ['Referência', 'Designação', 'Quant.', 'Pr.Unit.', 'Pr.Unit.Moeda', 'Tabela de IVA', 'Cor', 'Tamanho', 'TOTAL', 'Destino', 'CPO']
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='openpyxl') as writer:
                for aba_nom in df_final['Aba'].unique():
                    nome_s = str(aba_nom)[:31].replace('/', '-')
                    df_final[df_final['Aba'] == aba_nom][cols].to_excel(writer, sheet_name=nome_s, index=False)
            st.success(f"✅ Sucesso! Geradas {len(df_final)} linhas.")
            st.download_button("⬇️ Descarregar Excel", out.getvalue(), f"IMPORTAR_{cliente}.xlsx")
        else:
            st.warning("Dados não encontrados. Verifique se selecionou o cliente correto para este ficheiro.")

    except Exception as e:
        st.error(f"Erro: {e}")
