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
        df_final_phc = pd.DataFrame()
        
        if cliente == "Stussy":
            df_original = xl.parse(xl.sheet_names[0], header=None)
            dados = df_original.iloc[1:].copy()
            df_phc = pd.DataFrame()
            df_phc['Referência'] = ""
            df_phc['Designação'] = ""
            df_phc['Quant.'] = pd.to_numeric(dados[12], errors='coerce').fillna(0)
            df_phc['Pr.Unit.'] = pd.to_numeric(dados[17], errors='coerce').fillna(0)
            df_phc['Pr.Unit.Moeda'] = df_phc['Pr.Unit.']
            df_phc['Tabela de IVA'] = 4
            df_phc['Cor'] = dados[6]
            df_phc['Tamanho'] = dados[9]
            df_phc['Destino'] = dados[4]
            df_final_phc = df_phc

        elif cliente == "Supreme":
            lista_phc = []
            # Percorrer todas as abas exceto a "TOTAL"
            for nome_aba in xl.sheet_names:
                if "TOTAL" in nome_aba.upper():
                    continue
                
                df_aba = xl.parse(nome_aba, header=None)
                destino = str(df_aba.iloc[16, 0]) # Célula A17 (índice 16, 0)
                
                # 1. Capturar os tamanhos na linha 15 (índices J:P são 9 a 15)
                tamanhos = {}
                for col_idx in range(9, 16):
                    valor_tam = df_aba.iloc[14, col_idx] # Linha 15 é índice 14
                    if pd.notna(valor_tam) and str(valor_tam).strip() != "":
                        tamanhos[col_idx] = str(valor_tam).strip()

                # 2. Percorrer linhas a partir da 18 (índice 17) para Cores, Quantidades e Preços
                for i in range(17, len(df_aba)):
                    cor = df_aba.iloc[i, 6] # Coluna G é índice 6
                    preco = df_aba.iloc[i, 17] # Coluna R é índice 17
                    
                    if pd.isna(cor) or str(cor).strip() == "":
                        continue
                        
                    # Para cada cor, ver as quantidades nos tamanhos identificados
                    for col_idx, nome_tamanho in tamanhos.items():
                        qtd = df_aba.iloc[i, col_idx]
                        if pd.notna(qtd) and pd.to_numeric(qtd, errors='coerce', default=0) > 0:
                            lista_phc.append({
                                'Referência': "",
                                'Designação': "",
                                'Quant.': qtd,
                                'Pr.Unit.': preco,
                                'Pr.Unit.Moeda': preco,
                                'Tabela de IVA': 4,
                                'Cor': cor,
                                'Tamanho': nome_tamanho,
                                'Destino': destino
                            })
            df_final_phc = pd.DataFrame(lista_phc)

        # --- GERAÇÃO DO FICHEIRO FINAL ---
        if not df_final_phc.empty:
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                for dest in df_final_phc['Destino'].unique():
                    if pd.isna(dest) or str(dest).strip() == "": dest = "Sem_Destino"
                    df_temp = df_final_phc[df_final_phc['Destino'] == dest].copy()
                    df_temp['TOTAL'] = df_temp['Quant.'] * df_temp['Pr.Unit.']
                    df_temp['CPO'] = ""
                    # Limpar nome da aba (máximo 31 chars)
                    nome_sheet = str(dest)[:25].replace('/', '-')
                    df_temp.to_excel(writer, sheet_name=nome_sheet, index=False)
            
            st.success(f"✅ Ficheiro {cliente} processado com sucesso!")
            st.download_button(
                label="⬇️ Descarregar Excel para PHC",
                data=output.getvalue(),
                file_name=f"IMPORTAR_PHC_{cliente}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Não foram encontrados dados válidos para processar.")

    except Exception as e:
        st.error(f"Erro no processamento: {e}")
