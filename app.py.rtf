import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="ROVO - Conversor PHC", page_icon="📦")

st.title("📦 Conversor de Encomendas ROVO")
st.info("Instruções: Carregue o Excel da Stussy e obtenha o ficheiro pronto para o PHC.")

# Upload do ficheiro
arquivo = st.file_uploader("Submeter ficheiro Excel", type=["xlsx"])

if arquivo:
    try:
        # Lógica de posições que definimos (M, R, G, J, E)
        df_original = pd.read_excel(arquivo, header=None)
        dados = df_original.iloc[1:].copy() # Salta o cabeçalho
        
        df_phc = pd.DataFrame()
        df_phc['Referência'] = ""
        df_phc['Designação'] = ""
        df_phc['Quant.'] = pd.to_numeric(dados[12], errors='coerce').fillna(0) # Coluna M
        df_phc['Pr.Unit.'] = pd.to_numeric(dados[17], errors='coerce').fillna(0) # Coluna R
        df_phc['Pr.Unit.Moeda'] = df_phc['Pr.Unit.']
        df_phc['Tabela de IVA'] = 4
        df_phc['Cor'] = dados[6] # Coluna G
        df_phc['Tamanho'] = dados[9] # Coluna J
        df_phc['TOTAL'] = df_phc['Quant.'] * df_phc['Pr.Unit.']
        df_phc['CPO'] = ""
        df_phc['Destino'] = dados[4] # Coluna E
        
        po_number = str(dados[2].iloc[0]) # Coluna C

        # Criar Excel em memória
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for destino in df_phc['Destino'].unique():
                if pd.isna(destino) or str(destino).strip() == "": continue
                df_temp = df_phc[df_phc['Destino'] == destino]
                nome_aba = str(destino)[:25].replace('/', '-')
                df_temp.to_excel(writer, sheet_name=nome_aba, index=False)
        
        st.success(f"✅ PO {po_number} processada!")
        st.download_button(
            label="⬇️ Descarregar para PHC",
            data=output.getvalue(),
            file_name=f"IMPORTAR_PHC_PO_{po_number}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Erro no processamento: {e}")
