import streamlit as st
import pandas as pd
import io
import pdfplumber
import re

st.set_page_config(page_title="ROVO - Conversor Universal", page_icon="🚀")
st.title("🚀 ROVO Universal Converter")

cliente = st.sidebar.selectbox("Selecione o Cliente", ["Stussy", "Supreme", "Studio Nicholson"])
st.info(f"Modo: **{cliente}**")

formatos = ["xlsx", "pdf"] if cliente == "Studio Nicholson" else ["xlsx"]
arquivo = st.file_uploader(f"Carregar ficheiro", type=formatos)

if arquivo:
    try:
        lista_dados = []

        # --- LÓGICA EXCEL (STUSSY / SUPREME) - Sem alterações ---
        if arquivo.name.endswith('.xlsx'):
            xl = pd.ExcelFile(arquivo, engine='openpyxl')
            if cliente == "Stussy":
                df = xl.parse(xl.sheet_names[0], header=None)
                for i, row in df.iloc[1:].iterrows():
                    q, p = pd.to_numeric(row[12], errors='coerce'), pd.to_numeric(row[17], errors='coerce')
                    if q and q > 0:
                        lista_dados.append({'Referência': "", 'Designação': "", 'Quant.': q, 'Pr.Unit.': 0, 'Pr.Unit.Moeda': p, 'Tabela de IVA': 4, 'Cor': row[6], 'Tamanho': row[9], 'TOTAL': q*(p if p else 0), 'Destino': row[4], 'Aba': "Stussy_PO"})
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
                                    lista_dados.append({'Referência': "", 'Designação': "", 'Quant.': q, 'Pr.Unit.': 0, 'Pr.Unit.Moeda': p, 'Tabela de IVA': 4, 'Cor': df.iloc[i, 6], 'Tamanho': t_nom, 'TOTAL': q*(p if p else 0), 'Destino': dest, 'Aba': aba})

        # --- LÓGICA PDF (STUDIO NICHOLSON) - REFORMULADA ---
        elif arquivo.name.endswith('.pdf') and cliente == "Studio Nicholson":
            with pdfplumber.open(arquivo) as pdf:
                for page in pdf.pages:
                    texto_completo = page.extract_text()
                    linhas = texto_completo.split('\n')
                    
                    destino = "Ver PDF"
                    modelo_atual = ""
                    nomes_tams = ["UK4", "UK6", "UK8", "UK10", "UK12", "UK14"]
                    
                    # 1. Encontrar o Destino (Ship To)
                    ship_match = re.search(r"Ship To:\s*(.*)", texto_completo, re.IGNORECASE)
                    if ship_match:
                        destino = ship_match.group(1).split('\n')[0].strip()

                    for idx, linha in enumerate(linhas):
                        # 2. Identificar o Modelo (Designação)
                        # Procuramos a linha que contém o código do modelo (ex: SORIN SNW-1868)
                        if "SNW-" in linha or "SNM-" in linha:
                            # Filtramos para pegar apenas a parte do nome e código
                            partes_modelo = linha.split()
                            # Tenta pegar o nome (SORIN) e o código (SNW-1868)
                            modelo_atual = ""
                            for p in partes_modelo:
                                modelo_atual += p + " "
                                if "SNW-" in p or "SNM-" in p:
                                    break
                            modelo_atual = modelo_atual.strip()
                            continue

                        # 3. Identificar linha de Cores e Quantidades (contém o símbolo €)
                        if "€" in linha:
                            partes = linha.split()
                            
                            # A Cor costuma ser o primeiro elemento de texto da linha
                            # Vamos garantir que não apanhamos lixo (como "MICRO" ou "RIB")
                            cor_candidata = partes[0]
                            if cor_candidata.upper() in ["MICRO", "JERSEY", "RIB"]:
                                # Se a primeira palavra for micro rib, a cor deve ser a segunda
                                cor_candidata = partes[1] if len(partes) > 1 else partes[0]
                            
                            # Preço Moeda
                            p_moeda = 0
                            try:
                                # O preço costuma estar a seguir ao símbolo €
                                if "€" in partes:
                                    p_idx = partes.index("€")
                                    p_txt = partes[p_idx+1].replace(',','').strip()
                                    p_moeda = pd.to_numeric(p_txt, errors='coerce')
                                else:
                                    # Caso o € esteja colado ao número
                                    for p in partes:
                                        if "€" in p:
                                            p_txt = p.replace('€','').replace(',','').strip()
                                            p_moeda = pd.to_numeric(p_txt, errors='coerce')
                                            break
                            except: pass

                            # Quantidades: Apanhamos todos os números antes do preço/total
                            # No seu PDF os tamanhos vêm antes da Qty Total
                            numeros = [n for n in partes if n.replace('.','').isdigit()]
                            # A Nicholson tem: [Qtd1, Qtd2, ..., QtdTotal]
                            # Removemos o último número que é a soma (Qty)
                            qts_puras = numeros[:-1] if len(numeros) > 1 else numeros
                            
                            for i_q, val_q in enumerate(qts_puras):
                                q_num = pd.to_numeric(val_q, errors='coerce')
                                if q_num and q_num > 0 and i_q < len(nomes_tams):
                                    lista_dados.append({
                                        'Referência': "", 
                                        'Designação': modelo_atual, 
                                        'Quant.': q_num, 
                                        'Pr.Unit.': 0, 
                                        'Pr.Unit.Moeda': p_moeda, 
                                        'Tabela de IVA': 4, 
                                        'Cor': cor_candidata, 
                                        'Tamanho': nomes_tams[i_q], 
                                        'TOTAL': q_num * (p_moeda if p_moeda else 0), 
                                        'Destino': destino, 
                                        'Aba': "Nicholson_PO"
                                    })

        # --- EXPORTAÇÃO ---
        df_final = pd.DataFrame(lista_dados)
        if not df_final.empty:
            df_final['CPO'] = ""
            cols = ['Referência', 'Designação', 'Quant.', 'Pr.Unit.', 'Pr.Unit.Moeda', 'Tabela de IVA', 'Cor', 'Tamanho', 'TOTAL', 'Destino', 'CPO']
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='openpyxl') as writer:
                for aba_nom in df_final['Aba'].unique():
                    df_final[df_final['Aba'] == aba_nom][cols].to_excel(writer, sheet_name=str(aba_nom)[:31], index=False)
            st.success(f"✅ Conversão concluída!")
            st.download_button("⬇️ Descarregar Excel PHC", out.getvalue(), f"IMPORTAR_{cliente}.xlsx")
        else:
            st.warning("Dados não encontrados. Verifique se o PDF é o original da Studio Nicholson.")

    except Exception as e:
        st.error(f"Erro: {e}")
