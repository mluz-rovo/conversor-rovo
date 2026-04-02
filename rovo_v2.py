import streamlit as st
import pandas as pd
import io
import pdfplumber
import re

st.set_page_config(page_title="ROVO - Conversor Universal", page_icon="🚀", layout="wide")
st.sidebar.title("🚀 MENU ROVO")
cliente = st.sidebar.selectbox("Selecione o Cliente", ["Studio Nicholson", "Stussy", "Supreme"])
st.title(f"📦 Conversor: {cliente}")

arquivo = st.file_uploader("Submeter ficheiro", type=["xlsx", "pdf"])

if arquivo:
    try:
        lista_dados = []
        # --- MANTEMOS A LÓGICA EXCEL PARA STUSSY/SUPREME ---
        if arquivo.name.endswith('.xlsx'):
            xl = pd.ExcelFile(arquivo)
            if cliente == "Stussy":
                df = xl.parse(xl.sheet_names[0], header=None)
                for i, row in df.iloc[1:].iterrows():
                    q, p = pd.to_numeric(row[12], errors='coerce'), pd.to_numeric(row[17], errors='coerce')
                    if q and q > 0:
                        lista_dados.append({'Referência': "", 'Designação': "", 'Quant.': q, 'Pr.Unit.': 0, 'Pr.Unit.Moeda': p, 'Tabela de IVA': 4, 'Cor': row[6], 'Tamanho': row[9], 'TOTAL': q*(p if p else 0), 'Destino': row[4], 'CPO': ""})
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
                                if q and q > 0:
                                    lista_dados.append({'Referência': "", 'Designação': "", 'Quant.': q, 'Pr.Unit.': 0, 'Pr.Unit.Moeda': p, 'Tabela de IVA': 4, 'Cor': df.iloc[i, 6], 'Tamanho': t_nom, 'TOTAL': q*(p if p else 0), 'Destino': dest, 'CPO': ""})

        # --- LÓGICA PDF STUDIO NICHOLSON (GRELHA DUPLA) ---
        elif arquivo.name.endswith('.pdf') and cliente == "Studio Nicholson":
            with pdfplumber.open(arquivo) as pdf:
                # Definição das duas grelhas possíveis
                grelha_letras = ["XXS", "XS", "S", "M", "L", "XL", "XXL"]
                grelha_numeros = ["UK4 / IT36", "UK6 / IT38", "UK8 / IT40", "UK10 / IT42", "UK12 / IT44", "UK14 / IT46"]
                
                # Palavras a ignorar na cor (REMOVIDO 'BLACK' E 'PANNA' DAQUI)
                lixo = ["JERSEY", "MICRO", "RIB", "MERCERIZED", "COTTON", "BRANDED", "BOXY", "FIT", "T-SHIRT", "QTY", "OTY", "TOTAL", "COST", "SNW", "SNM", "LAY"]

                for page in pdf.pages:
                    texto = page.extract_text()
                    if not texto: continue
                    linhas = texto.split('\n')
                    
                    modelo_atual = ""
                    destino = "Ver PDF"
                    grelha_ativa = []

                    for idx, linha in enumerate(linhas):
                        linha_up = linha.upper()
                        
                        # 1. Detetar o Modelo (Designação)
                        if any(x in linha_up for x in ["SNW -", "SNM -", "SN -", "LAY "]):
                            modelo_atual = linha.strip()
                            # Verificar qual a grelha de tamanhos que vem logo a seguir
                            next_l = linhas[idx+1].upper() if idx+1 < len(linhas) else ""
                            if "UK4" in next_l or "IT36" in next_l:
                                grelha_ativa = grelha_numeros
                            else:
                                grelha_ativa = grelha_letras
                            continue

                        # 2. Detetar linha de dados (contém o símbolo €)
                        if "€" in linha:
                            partes = linha.split()
                            
                            # Extrair Preço Unitário
                            p_unit = 0
                            precos = re.findall(r"(\d+[\.,]\d{2})", linha)
                            if precos:
                                p_unit = float(precos[0].replace(',', ''))

                            # Extrair Cor (Primeira palavra que não seja lixo ou número)
                            cor_final = "Ver PDF"
                            for p in partes:
                                p_clean = p.upper().replace(',', '').replace('.', '')
                                if p_clean not in lixo and not p_clean.isdigit() and "€" not in p_clean and len(p_clean) > 2:
                                    cor_final = p
                                    break
                            
                            # Extrair Quantidades (Números inteiros)
                            nums_qtd = [n for n in partes if n.isdigit() and len(n) < 4]
                            # Remove o total da linha (último número)
                            qts_reais = nums_qtd[:-1] if len(nums_qtd) > 1 else nums_qtd

                            # Associar quantidades ao tamanho correto da grelha ativa
                            for i_q, val_q in enumerate(qts_reais):
                                if i_q < len(grelha_ativa):
                                    q_num = int(val_q)
                                    if q_num > 0:
                                        lista_dados.append({
                                            'Referência': "", 'Designação': modelo_atual, 'Quant.': q_num,
                                            'Pr.Unit.': p_unit, 'Pr.Unit.Moeda': 0, 'Tabela de IVA': 4,
                                            'Cor': cor_final, 'Tamanho': grelha_ativa[i_q],
                                            'TOTAL': q_num * p_unit, 'Destino': destino, 'CPO': ""
                                        })

        if lista_dados:
            df_final = pd.DataFrame(lista_dados)
            cols = ['Referência', 'Designação', 'Quant.', 'Pr.Unit.', 'Pr.Unit.Moeda', 'Tabela de IVA', 'Cor', 'Tamanho', 'TOTAL', 'Destino', 'CPO']
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='openpyxl') as writer:
                df_final[cols].to_excel(writer, index=False, sheet_name="PHC")
            st.success("✅ Ficheiro 100439 e outros convertidos!")
            st.download_button("⬇️ Descarregar Excel", out.getvalue(), "IMPORTAR_STUDIO.xlsx")
        else:
            st.warning("Não foram encontrados dados. Verifique o ficheiro.")
    except Exception as e:
        st.error(f"Erro: {e}")
