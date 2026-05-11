import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.set_page_config(page_title="ROVO - Financial Control", page_icon="💰", layout="wide")

st.title("💰 Financial Control - Resumo CPO/SPO")

st.info("📤 Carregue o seu Excel com os dados financeiros para gerar um resumo agrupado por CPO e SPO")

# ===========================================================================
# UPLOAD DO FICHEIRO
# ===========================================================================
uploaded_file = st.file_uploader("Upload seu Excel financeiro", type=["xlsx"], key="financial_upload")

if uploaded_file:
    try:
        # Lê todas as abas do ficheiro
        excel_file = pd.ExcelFile(uploaded_file, engine="openpyxl")
        
        st.subheader("📊 Abas encontradas")
        st.write(excel_file.sheet_names)
        
        # Se houver múltiplas abas, combina todas
        if len(excel_file.sheet_names) > 1:
            all_sheets = {}
            for sheet in excel_file.sheet_names:
                all_sheets[sheet] = pd.read_excel(uploaded_file, sheet_name=sheet)
            df = pd.concat([df.assign(Source_Sheet=sheet) for sheet, df in all_sheets.items()], ignore_index=True)
        else:
            df = pd.read_excel(uploaded_file, sheet_name=0)
        
        # ===========================================================================
        # PREVIEW DOS DADOS
        # ===========================================================================
        st.subheader("📋 Preview dos Dados")
        with st.expander("Ver todos os dados", expanded=False):
            st.dataframe(df, use_container_width=True)
        
        st.write(f"✅ Carregadas **{len(df)}** linhas de dados")
        
        # ===========================================================================
        # CONFIGURAÇÃO DE COLUNAS
        # ===========================================================================
        st.subheader("⚙️ Mapear Colunas")
        st.write("Selecione as colunas que correspondem aos seus dados:")
        
        cols_list = df.columns.tolist()
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            col_shipping = st.selectbox("Shipping Destination", cols_list)
        with col2:
            col_collection = st.selectbox("Collection", cols_list)
        with col3:
            col_client_po = st.selectbox("Client PO", cols_list)
        with col4:
            col_cpo = st.selectbox("CPO", cols_list)
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            col_value_orig = st.selectbox("Value Original Currency", cols_list)
        with col2:
            col_currency = st.selectbox("Currency", cols_list)
        with col3:
            col_value_eur = st.selectbox("Value in €", cols_list)
        with col4:
            col_spo = st.selectbox("SPO", cols_list)
        
        col1, col2, col3 = st.columns(3)
        with col1:
            col_ship_date = st.selectbox("Estimated Shipping Date", cols_list)
        with col2:
            col_margin = st.selectbox("Direct Margin", cols_list)
        with col3:
            col_qty = st.selectbox("QTY", cols_list)
        
        col_supplier = st.selectbox("Supplier", cols_list)
        
        # ===========================================================================
        # PROCESSAR
        # ===========================================================================
        if st.button("🔄 Processar Dados", type="primary"):
            try:
                df_proc = df.copy()
                
                # Converter para números
                numeric_cols = [col_value_orig, col_value_eur, col_margin, col_qty]
                for col in numeric_cols:
                    if col in df_proc.columns:
                        df_proc[col] = pd.to_numeric(df_proc[col], errors='coerce').fillna(0)
                
                # ===========================================================================
                # RESUMO POR CPO
                # ===========================================================================
                st.subheader("📋 Resumo por CPO")
                
                agg_dict = {
                    col_shipping: 'first',
                    col_collection: 'first',
                    col_client_po: 'first',
                    col_value_orig: 'sum',
                    col_currency: 'first',
                    col_value_eur: 'sum',
                    col_ship_date: 'first',
                    col_margin: 'sum',
                    col_qty: 'sum',
                    col_supplier: 'first'
                }
                
                # Filtrar apenas colunas que existem
                agg_dict = {k: v for k, v in agg_dict.items() if k in df_proc.columns}
                
                summary_cpo = df_proc.groupby(col_cpo, as_index=False).agg(agg_dict)
                
                # Renomear para nomes finais
                summary_cpo = summary_cpo.rename(columns={
                    col_cpo: 'CPO',
                    col_shipping: 'Shipping Destination',
                    col_collection: 'Collection',
                    col_client_po: 'Client PO',
                    col_value_orig: 'Value in Original Currency',
                    col_currency: 'Currency',
                    col_value_eur: 'Value in €',
                    col_ship_date: 'Estimated Shipping Date',
                    col_margin: 'Direct Margin',
                    col_qty: 'QTY',
                    col_supplier: 'Supplier'
                })
                
                st.dataframe(summary_cpo, use_container_width=True, hide_index=True)
                
                # ===========================================================================
                # RESUMO POR CPO/SPO
                # ===========================================================================
                st.subheader("📋 Resumo Detalhado (CPO/SPO)")
                
                summary_spo = df_proc.groupby([col_cpo, col_spo], as_index=False).agg(agg_dict)
                
                summary_spo = summary_spo.rename(columns={
                    col_cpo: 'CPO',
                    col_spo: 'SPO',
                    col_shipping: 'Shipping Destination',
                    col_collection: 'Collection',
                    col_client_po: 'Client PO',
                    col_value_orig: 'Value in Original Currency',
                    col_currency: 'Currency',
                    col_value_eur: 'Value in €',
                    col_ship_date: 'Estimated Shipping Date',
                    col_margin: 'Direct Margin',
                    col_qty: 'QTY',
                    col_supplier: 'Supplier'
                })
                
                st.dataframe(summary_spo, use_container_width=True, hide_index=True)
                
                # ===========================================================================
                # ESTATÍSTICAS
                # ===========================================================================
                st.subheader("📊 Estatísticas")
                
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Total CPOs", len(summary_cpo))
                with col2:
                    st.metric("Total SPOs", len(summary_spo))
                with col3:
                    total_value = df_proc[col_value_eur].sum()
                    st.metric("Valor Total (€)", f"€{total_value:,.2f}")
                with col4:
                    total_margin = df_proc[col_margin].sum()
                    st.metric("Margem Total (€)", f"€{total_margin:,.2f}")
                
                # ===========================================================================
                # EXPORT
                # ===========================================================================
                st.subheader("⬇️ Descarregar Relatório")
                
                out = io.BytesIO()
                with pd.ExcelWriter(out, engine="openpyxl") as writer:
                    summary_spo.to_excel(writer, index=False, sheet_name="Sheet1")
                
                filename = f"Financial_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                
                st.download_button(
                    "📥 Download Relatório Financeiro",
                    out.getvalue(),
                    filename,
                    key="dl_financial",
                    type="primary"
                )
                
                st.success("✅ Relatório gerado com sucesso!")
                
            except Exception as e:
                st.error(f"❌ Erro ao processar: {e}")
                st.exception(e)
    
    except Exception as e:
        st.error(f"❌ Erro ao ler ficheiro: {e}")
        st.write("💡 Certifique-se que:")
        st.write("- O ficheiro é Excel (.xlsx)")
        st.write("- Tem as colunas esperadas (CPO, SPO, etc)")
        st.write("- As colunas numéricas contêm apenas números")

# ===========================================================================
# HELP
# ===========================================================================
with st.expander("ℹ️ Como Usar"):
    st.markdown("""
    ### 📋 Passo a Passo
    
    1. **Prepare seu Excel**
       - Colunas: Shipping Destination, Collection, Client PO, CPO, SPO, etc.
       - Veja o exemplo acima (Screenshot) para referência
    
    2. **Faça Upload**
       - Click em "Browse files"
       - Selecione seu Excel
    
    3. **Mapear Colunas**
       - Selecione qual coluna é CPO, SPO, Valor, etc.
       - O sistema identifica automaticamente
    
    4. **Processar**
       - Click em "Processar Dados"
       - Veja os resumos (CPO e CPO/SPO)
    
    5. **Descarregar**
       - Excel com 3 abas:
         - **CPO Summary**: Dados agrupados por CPO
         - **CPO/SPO Detail**: Dados agrupados por CPO e SPO
         - **Raw Data**: Dados originais
    
    ### 📊 O que Você Obtém
    
    - **Total CPOs**: Quantos CPOs únicos tem
    - **Total SPOs**: Quantos SPOs únicos tem
    - **Valor Total (€)**: Soma de todos os valores
    - **Margem Total (€)**: Soma de todas as margens
    """)
