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
        
        if arquivo.name.endswith('.xlsx'):
            # ... (Lógica Excel Stussy/Supreme mantida)
            pass

        elif arquivo.name.endswith('.pdf') and cliente == "Studio Nicholson":
            with pdfplumber.open(arquivo) as pdf:
                # Grelhas de referência para comparação
                tams_ref = ["XXS", "XS", "S", "M", "L", "XL", "XXL", "UK4", "UK6", "UK8", "UK10", "UK12", "UK14"]
                lixo_cor = ["JERSEY", "MICRO", "RIB", "SHORT", "SCOOP", "SLEEVE", "NECK", "VEST", "HENLEY", "COTTON", "BRANDED", "BOXY", "FIT", "T-SHIRT", "QTY", "COST", "TOTAL"]

                for page in pdf.pages:
                    texto = page.extract_text()
                    if not texto: continue
                    linhas = texto.split('\n')
                    
                    modelo_atual = ""
                    destino_atual = "Ver PDF"
                    mapeamento_tamanhos = {} # Guardará {coluna_index: "Tamanho"}

                    for i, linha in enumerate(linhas):
                        linha_up = linha.upper()
                        
                        if "SHIP TO:" in linha_up and i + 1 < len(linhas):
                            destino_atual = linhas[i+1].strip() [cite: 31]
                        
                        # 1. Identificar o Modelo e a Linha de Cabeçalho (Tamanhos)
                        if any(x in linha_up for x in ["SNW -", "SNM -", "SN -", "LAY "]):
                            modelo_atual = re.sub(r"\b(QTY|COST|TOTAL|FIRSTMAKE)\b", "", linha, flags=re.I).strip() [cite: 31]
                            
                            # Procurar a linha de tamanhos imediatamente abaixo ou na mesma
                            idx_busca = i
                            while idx_busca < min(i + 3, len(linhas)):
                                l_busca = linhas[idx_busca].upper()
                                partes_h = l_busca.split()
                                # Criar mapa de posições: ex {2: "S", 3: "M"}
                                mapeamento_tamanhos = {}
                                for idx_p, p in enumerate(partes_h):
                                    for t in tams_ref:
                                        if t == p or (t in p and "/" in p): # Trata UK4 / IT36
                                            mapeamento_tamanhos[idx_p] = p
                                            break
                                if mapeamento_tamanhos: break
                                idx_busca += 1
                            continue

                        # 2. Capturar Dados (Linha com €)
                        if "€" in linha:
                            partes = linha.split()
                            
                            # Preço Unitário
                            precos = re.findall(r"(\d+[\.,]\d{2})", linha)
                            p_unit = float(precos[0].replace(',', '')) if precos else 0 [cite: 31]
                            
                            # Cor (Filtro)
                            cor_candidata = ""
                            for p in partes:
                                p_up = p.upper().replace(',', '').replace('.', '')
                                if p_up not in lixo_cor and not p_up.isdigit() and "€" not in p_up and len(p_up) > 2:
                                    cor_candidata = p
                                    break [cite: 31]
                            
                            # Quantidades baseadas na POSIÇÃO do cabeçalho
                            for col_idx, tamanho_nome in mapeamento_tamanhos.items():
                                if col_idx < len(partes):
                                    val_q = partes[col_idx]
                                    if val_q.isdigit():
                                        q_num = int(val_q)
                                        if q_num > 0:
                                            lista_dados.append({
                                                'Referência': "", 'Designação': modelo_atual, 'Quant.': q_num,
                                                'Pr.Unit.': p_unit, 'Pr.Unit.Moeda': 0, 'Tabela de IVA': 4,
                                                'Cor': cor_candidata, 'Tamanho': tamanho_nome,
                                                'TOTAL': q_num * p_unit, 'Destino': destino_atual, 'CPO': ""
                                            }) [cite: 31]

        if lista_dados:
            df = pd.DataFrame(lista_dados)
            cols = ['Referência', 'Designação', 'Quant.', 'Pr.Unit.', 'Pr.Unit.Moeda', 'Tabela de IVA', 'Cor', 'Tamanho', 'TOTAL', 'Destino', 'CPO']
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='openpyxl') as writer:
                df[cols].to_excel(writer, index=False, sheet_name="PHC")
            st.success("✅ Conversão Inteligente Concluída!")
            st.download_button("⬇️ Descarregar Excel", out.getvalue(), "IMPORTAR_PHC.xlsx") [cite: 31]
    except Exception as e:
        st.error(f"Erro: {e}")
