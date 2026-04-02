import streamlit as st
import pandas as pd
import io
import tabula

st.set_page_config(page_title="ROVO - Conversor Universal", page_icon="🚀")
st.title("🚀 ROVO Universal Converter")

# Seleção do Cliente
cliente = st.sidebar.selectbox("Selecione o Cliente", ["Stussy", "Supreme", "Studio Nicholson"])
st.info(f"Modo: **{cliente}**")

# Aceita PDF para Nicholson e XLSX para os outros
formato = ["xlsx"] if cliente != "Studio Nicholson" else ["pdf", "xlsx"]
arquivo = st.file_uploader(f"Submeter ficheiro ({', '.join(formato)})", type=formato)

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
                    if q > 0:
                        lista_dados.append({'Referência': "", 'Designação': "", 'Quant.': q, 'Pr.Unit.': 0, 'Pr.Unit.Moeda': p, 'Tabela de IVA': 4, 'Cor': row[6], 'Tamanho': row[9], 'TOTAL': q*(p if p else 0), 'Destino': row[4], 'Aba': "Encomenda"})

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
                                if q > 0:
                                    lista_dados.append({'Referência': "", 'Designação': "", 'Quant.': q, 'Pr.Unit.': 0, 'Pr.Unit.Moeda': p, 'Tabela de IVA': 4, 'Cor': df.iloc[i, 6], 'Tamanho': t_nom, 'TOTAL': q*(p if p else 0), 'Destino': dest, 'Aba': aba})

        # --- LÓGICA PDF (STUDIO NICHOLSON) ---
        elif arquivo.name.endswith('.pdf') and cliente == "Studio Nicholson":
            # Extrai tabelas do PDF
            paginas = tabula.read_pdf(arquivo, pages='all', multiple_tables=True, pandas_options={'header': None})
            
            for df_pg in paginas:
                destino = "Ver PDF"
                modelo = ""
                # Mapeamento colunas PDF (ajustar após teste real)
                # Normalmente em PDF: Cor=Col1, Tam_S=Col5, Preço=Col Última
                for i, row in df_pg.iterrows():
                    txt = " ".join(row.astype(str))
                    if "Ship To:" in txt: destino = txt.split("Ship To:")[-1].strip()
                    if "SNW -" in txt or "SNM -" in txt: modelo = str(row[0])
                    
                    if "€" in txt:
                        p = pd.to_numeric(str(row.iloc[-2]).replace('€','').replace(',','.').strip(), errors='coerce')
                        # Exemplo de 5 colunas de tamanhos (ajustar índices conforme o PDF lido)
                        for idx, tam_nom in enumerate(["UK4", "UK6", "UK8", "UK10", "UK12", "UK14"], start=4):
                            q = pd.to_numeric(row.iloc[idx], errors='coerce')
                            if q > 0:
                                lista_dados.append({'Referência': modelo, 'Designação': "", 'Quant.': q, 'Pr.Unit.': 0, 'Pr.Unit.Moeda': p, 'Tabela de IVA': 4, 'Cor': row.iloc[1], 'Tamanho': tam_nom, 'TOTAL': q*(p if p else 0), 'Destino': destino, 'Aba': " Nicholson_PO"})

        # --- EXPORTAÇÃO FINAL ---
        df_final = pd.DataFrame(lista_dados)
        if not df_final.empty:
            df_final['CPO'] = ""
            cols = ['Referência', 'Designação', 'Quant.', 'Pr.Unit.', 'Pr.Unit.Moeda', 'Tabela de IVA', 'Cor', 'Tamanho', 'TOTAL', 'Destino', 'CPO']
            
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='openpyxl') as writer:
                for aba_nom in df_final['Aba'].unique():
                    df_aba = df_final[df_final['Aba'] == aba_nom][cols]
                    df_aba.to_excel(writer, sheet_name=str(aba_nom)[:31], index=False)
            
            st.success(f"✅ Convertido com sucesso!")
            st.download_button("⬇️ Descarregar Excel PHC", out.getvalue(), f"IMPORTAR_{cliente}.xlsx")
        else:
            st.warning("Não foram detetados dados válidos. Verifique se o formato do ficheiro corresponde ao cliente.")

    except Exception as e:
        st.error(f"Erro técnico: {e}. Certifique-se que o requirements.txt inclui 'tabula-py'.")
