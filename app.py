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
        lista_phc = []
        
        if cliente == "Stussy":
            df_original = xl.parse(xl.sheet_names[0], header=None)
            dados = df_original.iloc[1:].copy()
            for i, row in dados.iterrows():
                lista_phc.append({
                    'Referência': "",
                    'Designação': "",
                    'Quant.': pd.to_numeric(row[12], errors='coerce'),
                    'Pr.Unit.': pd.to_numeric(row[17], errors='coerce'),
                    'Pr.Unit.Moeda': pd.to_numeric(row[17], errors='coerce'),
                    'Tabela de IVA': 4,
                    'Cor': row[6],
                    'Tamanho': row[9],
                    'Destino': row[4]
                })

        elif cliente == "Supreme":
            for nome_aba in xl.sheet_names:
                if "TOTAL" in nome_aba.upper():
                    continue
                
                df_aba = xl.parse(nome_aba, header=None)
                # Célula A17 (índice 16, coluna 0)
                destino = str(df_aba.iloc[16, 0]).strip() if len(df_aba) > 16 else nome_aba
                
                # 1. Identificar tamanhos na linha 15 (índice 14, colunas J a P -> 9 a 15)
                tamanhos_encontrados = {}
                for col_idx in range(9, 16):
                    val = df_aba.iloc[14, col_idx]
                    if pd.notna(val) and str(val).strip() != "":
                        tamanhos_encontrados[col_idx] = str(val).strip()

                # 2. Ler dados a partir da linha 18 (índice 17)
                for i in range(17, len(df_aba)):
                    cor = df_aba.iloc[i, 6]   # Coluna G (índice 6)
                    preco = df_aba.iloc[i, 17] # Coluna R (índice 17)
                    
                    if pd.isna(cor) or str(cor).strip() == "":
                        continue
                        
                    for col_idx, nome_tam in tamanhos_encontrados.items():
                        qtd = pd.to_numeric(df_aba.iloc[i, col_idx], errors='coerce')
                        if pd.notna(qtd) and qtd > 0:
                            lista_phc.append({
                                'Referência': "",
                                'Designação': "",
                                'Quant.': qtd,
                                'Pr.Unit.': preco,
                                'Pr.Unit.Moeda': preco,
                                'Tabela de IVA': 4,
                                'Cor': cor,
                                'Tamanho': nome_tam,
                                'Destino': destino
                            })

        df_final = pd.DataFrame(lista_phc)

        if not df_final.empty:
            # Limpeza final de números
            df_final['Quant.'] = df_final['Quant.'].fillna(0)
            df_final['Pr.Unit.'] = df_final['Pr.Unit.'].fillna(0)
            df_final['TOTAL'] = df_final['Quant.'] * df_final['Pr.Unit.']
            df_final['CPO'] = ""

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                for dest in df_final['Destino'].unique():
                    aba_nome = str(dest)[:25].replace('/', '-') if pd.notna(dest) else "Geral"
                    df_temp = df_final[df_final['Destino'] == dest]
                    df_temp.to_excel(writer, sheet_name=aba_nome, index=False)
            
            st.success(f"✅ Sucesso! Geradas {len(df_final)} linhas de encomenda.")
            st.download_button(
                label="⬇️ Descarregar Excel para PHC",
                data=output.getvalue(),
                file_name=f"IMPORTAR_{cliente}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Atenção: Não foram encontrados dados nas linhas/colunas especificadas.")

    except Exception as e:
        st.error(f"Erro no processamento: {e}")
