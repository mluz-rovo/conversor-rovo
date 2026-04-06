import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="ROVO - Conversor Multi-Cliente", page_icon="📦")
st.title("📦 Conversor de Encomendas ROVO")

cliente = st.sidebar.selectbox("Selecione o Cliente", ["Stussy", "Supreme"])
st.info(f"A processar formato para: **{cliente}**")

arquivo = st.file_uploader("Submeter ficheiro Excel", type=["xlsx"])

if arquivo:
    try:
        xl = pd.ExcelFile(arquivo, engine='openpyxl')
        lista_dados = []
        
        if cliente == "Stussy":
            df_original = xl.parse(xl.sheet_names[0], header=None)
            dados = df_original.iloc[1:].copy()
            for i, row in dados.iterrows():
                q = pd.to_numeric(row[12], errors='coerce')
                p_moeda = pd.to_numeric(row[17], errors='coerce')
                lista_dados.append({
                    'Referência': "", 'Designação': "", 'Quant.': q,
                    'Pr.Unit.': 0, # Preço unitário base a zero
                    'Pr.Unit.Moeda': p_moeda, 
                    'Tabela de IVA': 4, 
                    'Cor': row[6], 'Tamanho': row[9], 
                    'TOTAL': (q if q else 0) * (p_moeda if p_moeda else 0),
                    'Destino': row[4], 'Aba_Original': xl.sheet_names[0]
                })

        elif cliente == "Supreme":
            for nome_aba in xl.sheet_names:
                if "TOTAL" in nome_aba.upper(): continue
                
                df_aba = xl.parse(nome_aba, header=None)
                
                # Identificar tamanhos (Linha 15 -> índice 14, Colunas J:P -> 9:15)
                tamanhos_gps = {}
                for col_idx in range(9, 16):
                    if col_idx < len(df_aba.columns):
                        val_tam = df_aba.iloc[14, col_idx]
                        if pd.notna(val_tam) and str(val_tam).strip() != "":
                            tamanhos_gps[col_idx] = str(val_tam).strip()

                # Blocos de 14 linhas (Destinos em 17, 31, 45...)
                for start_row in range(16, len(df_aba), 14):
                    destino_bloco = str(df_aba.iloc[start_row, 0]).strip()
                    if pd.isna(destino_bloco) or destino_bloco == "" or destino_bloco == "nan": continue
                    
                    for i in range(start_row + 1, start_row + 13):
                        if i >= len(df_aba): break
                        cor = df_aba.iloc[i, 6]
                        p_moeda = pd.to_numeric(df_aba.iloc[i, 17], errors='coerce')
                        if pd.isna(cor) or str(cor).strip() == "": continue
                            
                        for col_idx, nome_tam in tamanhos_gps.items():
                            qtd = pd.to_numeric(df_aba.iloc[i, col_idx], errors='coerce')
                            if pd.notna(qtd) and qtd > 0:
                                lista_dados.append({
                                    'Referência': "", 'Designação': "", 'Quant.': qtd,
                                    'Pr.Unit.': 0, # Preço unitário base a zero
                                    'Pr.Unit.Moeda': p_moeda, 
                                    'Tabela de IVA': 4,
                                    'Cor': cor, 'Tamanho': nome_tam, 
                                    'TOTAL': qtd * (p_moeda if p_moeda else 0),
                                    'Destino': destino_bloco,
                                    'Aba_Original': nome_aba
                                })
                                
        df_final = pd.DataFrame(lista_dados)

        if not df_final.empty:
            df_final['CPO'] = ""
            # Reordenar colunas conforme solicitado anteriormente
            colunas_phc = ['Referência', 'Designação', 'Quant.', 'Pr.Unit.', 'Pr.Unit.Moeda', 
                           'Tabela de IVA', 'Cor', 'Tamanho', 'TOTAL', 'Destino', 'CPO']
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                if cliente == "Supreme":
                    for aba in df_final['Aba_Original'].unique():
                        df_aba_res = df_final[df_final['Aba_Original'] == aba][colunas_phc]
                        df_aba_res.to_excel(writer, sheet_name=str(aba)[:31], index=False)
                else:
                    for dest in df_final['Destino'].unique():
                        df_dest_res = df_final[df_final['Destino'] == dest][colunas_phc]
                        nome_s = str(dest)[:31].replace('/', '-')
                        df_dest_res.to_excel(writer, sheet_name=nome_s, index=False)
            
            st.success(f"✅ Processado! Preços em 'Pr.Unit.Moeda' e 'Pr.Unit.' a zero.")
            st.download_button(label="⬇️ Descarregar para PHC", data=output.getvalue(), 
                               file_name=f"IMPORTAR_{cliente}.xlsx", mime="application/vnd.ms-excel")
    except Exception as e:
        st.error(f"Erro: {e}")
