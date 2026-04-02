import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="ROVO - Conversor Multi-Cliente", page_icon="📦")
st.title("📦 Conversor de Encomendas ROVO")

cliente = st.sidebar.selectbox("Selecione o Cliente", ["Stussy", "Supreme", "Studio Nicholson"])
st.info(f"A processar formato para: **{cliente}**")

arquivo = st.file_uploader("Submeter ficheiro Excel (Nicholson converter PDF para Excel primeiro)", type=["xlsx"])

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
                    'Referência': "", 'Designação': "", 'Quant.': q, 'Pr.Unit.': 0,
                    'Pr.Unit.Moeda': p_moeda, 'Tabela de IVA': 4, 'Cor': row[6], 'Tamanho': row[9],
                    'TOTAL': (q if q else 0) * (p_moeda if p_moeda else 0), 'Destino': row[4], 'Aba_Original': xl.sheet_names[0]
                })

        elif cliente == "Supreme":
            for nome_aba in xl.sheet_names:
                if "TOTAL" in nome_aba.upper(): continue
                df_aba = xl.parse(nome_aba, header=None)
                tamanhos_gps = {}
                for col_idx in range(9, 16):
                    if col_idx < len(df_aba.columns):
                        val_tam = df_aba.iloc[14, col_idx]
                        if pd.notna(val_tam) and str(val_tam).strip() != "": tamanhos_gps[col_idx] = str(val_tam).strip()

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
                                    'Referência': "", 'Designação': "", 'Quant.': qtd, 'Pr.Unit.': 0,
                                    'Pr.Unit.Moeda': p_moeda, 'Tabela de IVA': 4, 'Cor': cor, 'Tamanho': nome_tam,
                                    'TOTAL': qtd * (p_moeda if p_moeda else 0), 'Destino': destino_bloco, 'Aba_Original': nome_aba
                                })

        elif cliente == "Studio Nicholson":
            for nome_aba in xl.sheet_names:
                df_aba = xl.parse(nome_aba, header=None)
                
                # 1. Tentar encontrar o destino "Ship To:"
                destino = "Desconhecido"
                for r in range(min(30, len(df_aba))):
                    linha_completa = " ".join(df_aba.iloc[r, :].astype(str))
                    if "Ship To:" in linha_completa:
                        # O destino costuma estar algumas linhas abaixo ou colunas ao lado
                        destino = str(df_aba.iloc[r+1, 2]).strip()
                        break

                modelo_atual = ""
                # Mapeamento de colunas para Nicholson (ajustar se o conversor PDF mudar as colunas)
                # No print: UK4=Col J(9), UK6=Col K(10), etc.
                indices_tamanhos = {9: "UK4", 10: "UK6", 11: "UK8", 12: "UK10", 13: "UK12", 14: "UK14"}

                for i in range(len(df_aba)):
                    linha = df_aba.iloc[i, :].astype(str).tolist()
                    linha_str = " ".join(linha)

                    # Detetar Modelo (Ex: SORIN SNW - 1868)
                    if "SNW -" in linha_str or "SNM -" in linha_str:
                        modelo_atual = str(df_aba.iloc[i, 0]).strip()
                        continue

                    # Detetar linha de dados (contém o símbolo do Euro e Cor na Coluna E/4)
                    if "€" in linha_str:
                        cor = str(df_aba.iloc[i, 4]).strip()
                        # Preço está na coluna 18 (S) ou 17 (R) no print "First/Make Cost"
                        preco_bruto = str(df_aba.iloc[i, 18]).replace('€','').replace(',','.').strip()
                        p_moeda = pd.to_numeric(preco_bruto, errors='coerce')
                        
                        for col_idx, nome_tam in indices_tamanhos.items():
                            if col_idx < len(df_aba.columns):
                                qtd = pd.to_numeric(df_aba.iloc[i, col_idx], errors='coerce')
                                if pd.notna(qtd) and qtd > 0:
                                    lista_dados.append({
                                        'Referência': modelo_atual,
                                        'Designação': "", 'Quant.': qtd, 'Pr.Unit.': 0,
                                        'Pr.Unit.Moeda': p_moeda, 'Tabela de IVA': 4,
                                        'Cor': cor, 'Tamanho': nome_tam,
                                        'TOTAL': qtd * (p_moeda if p_moeda else 0),
                                        'Destino': destino, 'Aba_Original': nome_aba
                                    })

        df_final = pd.DataFrame(lista_dados)
        if not df_final.empty:
            df_final['CPO'] = ""
            colunas_phc = ['Referência', 'Designação', 'Quant.', 'Pr.Unit.', 'Pr.Unit.Moeda', 
                           'Tabela de IVA', 'Cor', 'Tamanho', 'TOTAL', 'Destino', 'CPO']
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # Agrupamento por Aba para Nicholson e Supreme
                agrupar_por = 'Aba_Original' if cliente in ["Supreme", "Studio Nicholson"] else 'Destino'
                for grupo in df_final[agrupar_por].unique():
                    df_res = df_final[df_final[agrupar_por] == grupo][colunas_phc]
                    nome_aba_final = str(grupo)[:31].replace('/', '-')
                    df_res.to_excel(writer, sheet_name=nome_aba_final, index=False)
            
            st.success(f"✅ Processamento de {cliente} concluído!")
            st.download_button(label="⬇️ Descarregar para PHC", data=output.getvalue(), 
                               file_name=f"IMPORTAR_{cliente}.xlsx", mime="application/vnd.ms-excel")
    except Exception as e:
        st.error(f"Erro no processamento: {e}")
