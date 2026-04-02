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
        
        if cliente == "Stussy":
            # (Mantemos a lógica da Stussy igual porque já funcionava)
            df_original = xl.parse(xl.sheet_names[0], header=None)
            dados = df_original.iloc[1:].copy()
            lista_stussy = []
            for i, row in dados.iterrows():
                lista_stussy.append({
                    'Referência': "", 'Designação': "", 'Quant.': pd.to_numeric(row[12], errors='coerce'),
                    'Pr.Unit.': pd.to_numeric(row[17], errors='coerce'), 'Pr.Unit.Moeda': pd.to_numeric(row[17], errors='coerce'),
                    'Tabela de IVA': 4, 'Cor': row[6], 'Tamanho': row[9], 'Destino': row[4]
                })
            df_final = pd.DataFrame(lista_stussy)

        elif cliente == "Supreme":
            lista_supreme = []
            for nome_aba in xl.sheet_names:
                if "TOTAL" in nome_aba.upper():
                    continue
                
                df_aba = xl.parse(nome_aba, header=None)
                
                # Identificar tamanhos (Linha 15 -> índice 14, Colunas J:P -> 9:15)
                tamanhos_gps = {}
                for col_idx in range(9, 16):
                    val_tam = df_aba.iloc[14, col_idx]
                    if pd.notna(val_tam) and str(val_tam).strip() != "":
                        tamanhos_gps[col_idx] = str(val_tam).strip()

                # Percorrer a aba em blocos de 14 linhas (Destinos em 17, 31, 45...)
                # Começamos na linha 17 (índice 16)
                for start_row in range(16, len(df_aba), 14):
                    # O nome do destino está na Coluna A desta linha
                    novo_destino = str(df_aba.iloc[start_row, 0]).strip()
                    if pd.isna(novo_destino) or novo_destino == "" or novo_destino == "nan":
                        continue
                    
                    # Os dados de cores/quantidades começam uma linha abaixo do destino
                    # E duram cerca de 12 linhas antes do próximo destino
                    for i in range(start_row + 1, start_row + 13):
                        if i >= len(df_aba): break
                        
                        cor = df_aba.iloc[i, 6]   # Coluna G (índice 6)
                        preco = df_aba.iloc[i, 17] # Coluna R (índice 17)
                        
                        if pd.isna(cor) or str(cor).strip() == "":
                            continue
                            
                        for col_idx, nome_tam in tamanhos_gps.items():
                            qtd = pd.to_numeric(df_aba.iloc[i, col_idx], errors='coerce')
                            if pd.notna(qtd) and qtd > 0:
                                lista_supreme.append({
                                    'Referência': "", 'Designação': "", 'Quant.': qtd,
                                    'Pr.Unit.': preco, 'Pr.Unit.Moeda': preco, 'Tabela de IVA': 4,
                                    'Cor': cor, 'Tamanho': nome_tam, 
                                    'Destino': novo_destino, # Local específico do bloco
                                    'Aba_Original': nome_aba # Para mantermos o nome da aba
                                })
            df_final = pd.DataFrame(lista_supreme)

        if not df_final.empty:
            df_final['TOTAL'] = df_final['Quant.'] * df_final['Pr.Unit.'].fillna(0)
            df_final['CPO'] = ""

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # Para a Supreme, agrupamos pelo nome da Aba Original
                if cliente == "Supreme":
                    for aba in df_final['Aba_Original'].unique():
                        df_aba_res = df_final[df_final['Aba_Original'] == aba].drop(columns=['Aba_Original'])
                        df_aba_res.to_excel(writer, sheet_name=str(aba)[:31], index=False)
                else:
                    for dest in df_final['Destino'].unique():
                        nome_s = str(dest)[:31].replace('/', '-')
                        df_final[df_final['Destino'] == dest].to_excel(writer, sheet_name=nome_s, index=False)
            
            st.success(f"✅ Processamento concluído!")
            st.download_button(label="⬇️ Descarregar para PHC", data=output.getvalue(), 
                               file_name=f"IMPORTAR_{cliente}.xlsx", mime="application/vnd.ms-excel")
        else:
            st.warning("Nenhum dado encontrado.")

    except Exception as e:
        st.error(f"Erro: {e}")
