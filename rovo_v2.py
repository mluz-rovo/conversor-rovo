import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.set_page_config(page_title="ROVO - Financial Control", page_icon="💰", layout="wide")

st.title("💰 Financial Control - Resumo CPO/SPO")

st.info("📤 Carregue o seu Excel. As colunas serão detectadas automaticamente e agrupadas por CPO/SPO.")

uploaded_file = st.file_uploader("Upload seu Excel financeiro", type=["xlsx"], key="financial_upload")

if uploaded_file:
    try:
        # Lê o Excel
        excel_file = pd.ExcelFile(uploaded_file, engine="openpyxl")
        
        # Se múltiplas abas, combina
        if len(excel_file.sheet_names) > 1:
            all_sheets = {sheet: pd.read_excel(uploaded_file, sheet_name=sheet) for sheet in excel_file.sheet_names}
            df_combined = pd.concat([df.assign(Source_Sheet=sheet) for sheet, df in all_sheets.items()], ignore_index=True)
        else:
            df_combined = pd.read_excel(uploaded_file, sheet_name=0)
        
        st.subheader("📋 Preview dos Dados")
        st.write(f"✅ Carregadas **{len(df_combined)}** linhas")
        with st.expander("Ver dados completos", expanded=False):
            st.dataframe(df_combined, use_container_width=True)
        
        # ===== DETECTAR COLUNAS =====
        cols_disponíveis = df_combined.columns.tolist()
        st.subheader("⚙️ Colunas Detectadas")
        st.write(cols_disponíveis)
        
        # Procurar por cada coluna
        col_cpo = None
        col_spo = None
        col_shipping = None
        col_collection = None
        col_client_po = None
        col_value_orig = None
        col_currency = None
        col_value_eur = None
        col_ship_date = None
        col_margin = None
        col_qty = None
        col_supplier = None
        
        for col in cols_disponíveis:
            col_lower = col.lower()
            if 'cpo' in col_lower and col_cpo is None:
                col_cpo = col
            if 'spo' in col_lower and col_spo is None:
                col_spo = col
            if ('shipping' in col_lower or 'destination' in col_lower) and col_shipping is None:
                col_shipping = col
            if 'collection' in col_lower and col_collection is None:
                col_collection = col
            if 'client' in col_lower and 'po' in col_lower and col_client_po is None:
                col_client_po = col
            if 'value in original' in col_lower and col_value_orig is None:
                col_value_orig = col
            if 'currency' in col_lower and col_currency is None:
                col_currency = col
            if ('value in €' in col_lower or 'value_eur' in col_lower) and col_value_eur is None:
                col_value_eur = col
            if ('estimated' in col_lower or 'shipping date' in col_lower) and 'date' in col_lower and col_ship_date is None:
                col_ship_date = col
            if ('margin' in col_lower or 'direct margin' in col_lower) and col_margin is None:
                col_margin = col
            if ('qty' in col_lower or 'quant' in col_lower) and col_qty is None:
                col_qty = col
            if 'supplier' in col_lower and col_supplier is None:
                col_supplier = col
        
        # Verificar se encontrou
        if not col_cpo:
            st.error("❌ Coluna CPO não encontrada!")
            st.write(f"Colunas disponíveis: {cols_disponíveis}")
        elif not col_spo:
            st.error("❌ Coluna SPO não encontrada!")
        else:
            st.success(f"✅ CPO: {col_cpo}, SPO: {col_spo}")
            
            if st.button("🔄 Processar Dados", type="primary"):
                try:
                    df_proc = df_combined.copy()
                    
                    # Converter para números
                    numeric_cols = [col_value_orig, col_value_eur, col_margin, col_qty]
                    for col in numeric_cols:
                        if col and col in df_proc.columns:
                            df_proc[col] = pd.to_numeric(df_proc[col], errors='coerce').fillna(0)
                    
                    # ===== AGRUPAR =====
                    agg_dict = {}
                    
                    # Números (soma)
                    numeric_agg = [col_value_orig, col_value_eur, col_margin, col_qty]
                    for col in numeric_agg:
                        if col and col in df_proc.columns:
                            agg_dict[col] = 'sum'
                    
                    # Todas as outras colunas (primeiro valor)
                    for col in df_proc.columns:
                        if col not in agg_dict and col != col_cpo and col != col_spo:
                            agg_dict[col] = 'first'
                    
                    # Agrupar por CPO e SPO
                    summary = df_proc.groupby([col_cpo, col_spo], as_index=False).agg(agg_dict)
                    
                    # Reordenar colunas (colocar CPO e SPO primeiro)
                    cols_ordem = [col_cpo, col_spo]
                    for col in cols_disponíveis:
                        if col not in cols_ordem and col in summary.columns:
                            cols_ordem.append(col)
                    
                    summary = summary[cols_ordem]
                    
                    st.subheader(f"📋 Resumo CPO/SPO ({len(summary)} linhas)")
                    st.dataframe(summary, use_container_width=True, hide_index=True)
                    
                    # ===== ESTATÍSTICAS =====
                    st.subheader("📊 Estatísticas")
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("Total CPOs", summary[col_cpo].nunique())
                    with col2:
                        st.metric("Total SPOs", len(summary))
                    with col3:
                        if col_value_eur and col_value_eur in summary.columns:
                            total = summary[col_value_eur].sum()
                            st.metric("Valor Total (€)", f"€{total:,.2f}")
                    with col4:
                        if col_margin and col_margin in summary.columns:
                            total = summary[col_margin].sum()
                            st.metric("Margem Total (€)", f"€{total:,.2f}")
                    
                    # ===== EXPORT =====
                    st.subheader("⬇️ Descarregar")
                    
                    out = io.BytesIO()
                    with pd.ExcelWriter(out, engine="openpyxl") as writer:
                        summary.to_excel(writer, index=False, sheet_name="Sheet1", startrow=0)
                        
                        # Adicionar fórmula se possível
                        if col_value_eur and col_currency and col_value_orig:
                            try:
                                ws = writer.sheets["Sheet1"]
                                
                                # Encontrar índices das colunas
                                col_idx_eur = list(summary.columns).index(col_value_eur) + 1
                                col_idx_currency = list(summary.columns).index(col_currency) + 1
                                col_idx_orig = list(summary.columns).index(col_value_orig) + 1
                                
                                # Adicionar fórmula para cada linha
                                for row_idx in range(2, len(summary) + 2):
                                    currency_cell = f"{chr(64 + col_idx_currency)}{row_idx}"
                                    value_orig_cell = f"{chr(64 + col_idx_orig)}{row_idx}"
                                    formula = f'=IF({currency_cell}="EUR",{value_orig_cell},{value_orig_cell}/$F$1)'
                                    ws[f"{chr(64 + col_idx_eur)}{row_idx}"] = formula
                                st.info("✅ Fórmula de câmbio adicionada na coluna 'Value in €'")
                            except Exception as e:
                                st.warning(f"Não conseguiu adicionar fórmula: {e}")
                    
                    filename = f"Financial_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                    st.download_button("📥 Download Relatório", out.getvalue(), filename, type="primary")
                    st.success("✅ Pronto!")
                    
                except Exception as e:
                    st.error(f"❌ Erro: {e}")
                    import traceback
                    st.write(traceback.format_exc())
    
    except Exception as e:
        st.error(f"❌ Erro ao ler ficheiro: {e}")
