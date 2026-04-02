import streamlit as st
import pandas as pd
import io
import tabula

st.set_page_config(page_title="ROVO - Conversor Universal", page_icon="🚀")
st.title("🚀 ROVO Universal Converter")

# Seleção do Cliente no Menu Lateral
cliente = st.sidebar.selectbox("Selecione o Cliente", ["Stussy", "Supreme", "Studio Nicholson"])
st.sidebar.markdown("---")
st.sidebar.write("**Instruções:**")
st.sidebar.write("1. Selecione o cliente.")
st.sidebar.write("2. Submeta o ficheiro (XLSX para Stussy/Supreme, PDF para Nicholson).")

st.info(f"Modo de Processamento: **{cliente}**")

# Formatos aceites conforme o cliente
formatos = ["xlsx", "pdf"] if cliente == "Studio Nicholson" else ["xlsx"]
arquivo = st.file_uploader(f"Carregar ficheiro do cliente", type=formatos)

if arquivo:
    try:
        lista_dados = []

        # --- CASO 1: EXCEL (Stussy ou Supreme) ---
        if arquivo.name.endswith('.xlsx'):
            xl = pd.ExcelFile(arquivo, engine='openpyxl')
            
            if cliente == "Stussy":
                df = xl.parse(xl.sheet_names[0], header=None)
                for i, row in df.iloc[1:].iterrows():
                    q = pd.to_numeric(row[12], errors='coerce')
                    p = pd.to_numeric(row[17], errors='coerce')
                    if q and q > 0:
                        lista_dados.append({
                            'Referência': "", 'Designação': "", 'Quant.': q, 'Pr.Unit.': 0, 
                            'Pr.Unit.Moeda': p, 'Tabela de IVA': 4, 'Cor': row[6], 'Tamanho': row[9], 
                            'TOTAL': q * (p if p else 0), 'Destino': row[4], 'Aba': "Stussy_PO"
                        })

            elif cliente == "Supreme":
                for aba in xl.sheet_names:
                    if "TOTAL" in aba.upper(): continue
                    df = xl.parse(aba, header=None)
                    # Grelha de tamanhos na linha 15 (J:P)
                    tams = {c: str(df.iloc[14, c]) for c in range(9, 16) if pd.notna(df.iloc[14, c])}
                    # Blocos de 14 linhas com destinos
                    for start in range(16, len(df), 14):
                        dest = str(df.iloc[start, 0]).strip()
                        for i in range(start + 1, start + 13):
                            if i >= len(df) or pd.isna(df.iloc[i, 6]): continue
                            p = pd.to_numeric(df.iloc[i, 17], errors='coerce')
                            for c_idx, t_nom in tams.items():
                                q = pd.to_numeric(df.iloc[i, c_idx], errors='coerce')
                                if q and q > 0:
                                    lista_dados.append({
                                        'Referência': "", 'Designação': "", 'Quant.': q, 'Pr.Unit.': 0, 
                                        'Pr.Unit.Moeda': p, 'Tabela de IVA': 4, 'Cor': df.iloc[i, 6], 
                                        'Tamanho': t_nom, 'TOTAL': q * (p if p else 0), 'Destino': dest, 'Aba': aba
                                    })

        # --- CASO 2: PDF (Studio Nicholson) ---
        elif arquivo.name.endswith('.pdf') and cliente == "Studio Nicholson":
            # Extrair tabelas do PDF usando Tabula
            tabelas_pdf = tabula.read_pdf(arquivo, pages='all', multiple_tables=True, pandas_options={'header': None})
            
            for df_pg in tabelas_pdf:
                destino_atual = "Ver no PDF"
                modelo_atual = "N/A"
                
                for i, row in df_pg.iterrows():
                    linha_txt = " ".join(row.astype(str))
                    
                    # Detetar Destino (Ship To:)
                    if "Ship To:" in linha_txt:
                        destino_atual = linha_txt.split("Ship To:")[-1].strip()
                    
                    # Detetar Modelo/Referência (SNW ou SNM)
                    if "SNW -" in linha_txt or "SNM -" in linha_txt:
                        modelo_atual = str(row[0]).strip()
                    
                    # Detetar Linha de Produção (tem o preço em €)
                    if "€" in linha_txt:
                        cor = str(row.iloc[1]).strip() # Coluna da Cor no PDF costuma ser a 2ª
                        # Preço costuma estar na penúltima coluna
                        p_texto = str(row.iloc[-2]).replace('€','').replace(',','.').strip()
                        p_moeda = pd.to_numeric(p_texto, errors='coerce')
                        
                        # Mapeamento de colunas de tamanhos (UK4 a UK14)
                        # Nota: Em PDF isto pode variar, mas tentamos as colunas do meio
                        for idx, tam_nom in enumerate(["UK4", "UK6", "UK8", "UK10", "UK12", "UK14"], start=4):
                            if idx < len(row):
                                q = pd.to_numeric(row.iloc[idx], errors='coerce')
                                if q and q > 0:
                                    lista_dados.append({
                                        'Referência': modelo_atual, 'Designação': "", 'Quant.': q, 'Pr.Unit.': 0, 
                                        'Pr.Unit.Moeda': p_moeda, 'Tabela de IVA': 4, 'Cor': cor, 
                                        'Tamanho': tam_nom, 'TOTAL': q * (p_moeda if p_moeda else 0), 
                                        'Destino': destino_atual, 'Aba': "Nicholson_PO"
                                    })

        # --- FINALIZAÇÃO E DOWNLOAD ---
        df_final = pd.DataFrame(lista_dados)
        if not df_final.empty:
            df_final['CPO'] = ""
            ordem_colunas = ['Referência', 'Designação', 'Quant.', 'Pr.Unit.', 'Pr.Unit.Moeda', 
                            'Tabela de IVA', 'Cor', 'Tamanho', 'TOTAL', 'Destino', 'CPO']
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # Criar abas conforme a origem
                agrupar = 'Aba'
                for nome_aba in df_final[agrupar].unique():
                    df_aba = df_final[df_final[agrupar] == nome_aba][ordem_colunas]
                    df_aba.to_excel(writer, sheet_name=str(nome_aba)[:31], index=False)
            
            st.success(f"✅ Conversão concluída!")
            st.download_button(label="⬇️ Descarregar Excel para PHC", 
                               data=output.getvalue(), 
                               file_name=f"IMPORTAR_{cliente}.xlsx", 
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.warning("Não foram encontrados dados. Verifique se o ficheiro e o cliente selecionado coincidem.")

    except Exception as e:
        st.error(f"Ocorreu um erro: {e}")
        st.info("Dica: Se o erro for 'tabula', verifique se o requirements.txt está correto.")
