import streamlit as st
import pandas as pd
import pdfplumber
import re
import io

# ── Padrões ────────────────────────────────────────────────────────────────────

# Bloco completo de produto:
# SORIN SNW - 1868 MICRO RIB  [sizes]  JERSEY - SHORT SLEEVE HENLEY BLACK  11 9 8 6 3 -  37  € 29.45  € total
PRODUCT_PATTERN = re.compile(
    r"([A-Z][\w\s\-]+?SNW\s*-\s*\d+[\w\s\-]*?)"          # 1: modelo  ex: "SORIN SNW - 1868 MICRO RIB"
    r"((?:UK\d+\s*/\s*IT\d+\s*)+)"                         # 2: tamanhos ex: "UK4 / IT36UK6 / IT38..."
    r"(JERSEY\s*-\s*[\w\s]+?)\s+"                          # 3: descrição+cor ex: "JERSEY - SHORT SLEEVE HENLEY BLACK"
    r"((?:\d+|-)\s+(?:(?:\d+|-)\s+)*)"                     # 4: quantidades ex: "11 9 8 6 3 - "
    r"(\d+)\s+"                                             # 5: qty total
    r"€\s*([\d,\.]+)\s+"                                   # 6: preço unitário
    r"€\s*[\d,\.]+",                                        # total (ignorado)
    re.IGNORECASE
)

SIZE_RE  = re.compile(r"UK(\d+)\s*/\s*IT\d+", re.IGNORECASE)
SHIP_RE  = re.compile(r"Ship To:\s*(.+?)(?=Bunschotenweg|Rua |Street|Avenue|Road|Docket Number)", re.IGNORECASE | re.DOTALL)


def extract_ship_to(text):
    m = re.search(r"Ship To:\s*([\w\s\-]+?)(?=Bunschotenweg|Rua |Street|Avenue|Road|Docket)", text, re.IGNORECASE)
    if m:
        return m.group(1).strip()
    return "Ver PDF"


def parse_modelo(texto):
    """Extrai só 'NOME SNW-XXXX' do modelo completo."""
    m = re.search(r"([A-Z]+\s+SNW\s*-\s*\d+)", texto, re.IGNORECASE)
    if m:
        return re.sub(r"\s*-\s*", "-", m.group(1).strip())
    return texto.strip()


def parse_color(descricao):
    """Última(s) palavra(s) após a descrição do artigo são a cor."""
    # Remove prefixo "JERSEY - TIPO SUBTIPO" e fica com a cor
    partes = descricao.strip().split()
    noise = {"JERSEY", "-", "SHORT", "SLEEVE", "HENLEY", "SCOOP", "NECK",
             "VEST", "BOXY", "FIT", "T-SHIRT", "COTTON", "BRANDED", "CREW"}
    cor = [p for p in partes if p.upper() not in noise]
    return " ".join(cor) if cor else descricao.strip()


def extract_studio_nicholson(pdf_file):
    rows = []

    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            texto = page.extract_text() or ""
            destino = extract_ship_to(texto)

            for m in PRODUCT_PATTERN.finditer(texto):
                modelo     = parse_modelo(m.group(1))
                sizes_raw  = m.group(2)
                desc_cor   = m.group(3).strip()
                qtys_raw   = m.group(4).strip().split()
                preco_str  = m.group(6).replace(",", "")
                preco      = float(preco_str)
                cor        = parse_color(desc_cor)

                # Extrair tamanhos mantendo ordem
                tamanhos = [f"UK{s}/IT{t}" for s, t in
                            re.findall(r"UK(\d+)\s*/\s*IT(\d+)", sizes_raw, re.IGNORECASE)]

                # Quantidades (dígito ou "-")
                quantidades = [int(q) if q.isdigit() else 0 for q in qtys_raw]

                for idx, tam in enumerate(tamanhos):
                    qty = quantidades[idx] if idx < len(quantidades) else 0
                    if qty > 0:
                        rows.append({
                            "Referência":    "",
                            "Designação":    modelo,
                            "Quant.":        qty,
                            "Pr.Unit.":      preco,
                            "Pr.Unit.Moeda": 0,
                            "Tabela de IVA": 4,
                            "Cor":           cor,
                            "Tamanho":       tam,
                            "TOTAL":         round(qty * preco, 2),
                            "Destino":       destino,
                            "CPO":           "",
                        })
    return rows


# ── Streamlit UI ───────────────────────────────────────────────────────────────

st.set_page_config(page_title="ROVO - Conversor Universal", page_icon="🚀", layout="wide")
st.sidebar.title("🚀 MENU ROVO")
cliente = st.sidebar.selectbox("Selecione o Cliente", ["Studio Nicholson", "Stussy", "Supreme"])
st.title(f"📦 Conversor: {cliente}")

arquivo = st.file_uploader("Submeter ficheiro", type=["xlsx", "pdf"])

if arquivo:
    try:
        lista_dados = []

        if arquivo.name.endswith(".xlsx"):
            pass  # Lógica Stussy/Supreme mantida aqui — não alterada

        elif arquivo.name.endswith(".pdf") and cliente == "Studio Nicholson":
            lista_dados = extract_studio_nicholson(arquivo)

        if lista_dados:
            cols = ["Referência", "Designação", "Quant.", "Pr.Unit.", "Pr.Unit.Moeda",
                    "Tabela de IVA", "Cor", "Tamanho", "TOTAL", "Destino", "CPO"]
            df  = pd.DataFrame(lista_dados).drop_duplicates()
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine="openpyxl") as writer:
                df[cols].to_excel(writer, index=False, sheet_name="PHC")
            out.seek(0)

            st.success(f"✅ {len(df)} linhas extraídas com sucesso!")
            st.dataframe(df[cols], use_container_width=True)
            st.download_button("⬇️ Download Excel", out.getvalue(), "IMPORTAR_PHC.xlsx")
        elif arquivo:
            st.warning("⚠️ Nenhum dado extraído. Verifica o PDF ou o cliente selecionado.")

    except Exception as e:
        st.error(f"Erro: {e}")
        st.exception(e)

def extract_ship_to(text):
    # DEBUG — remove depois
    idx = text.find("Ship To:")
    if idx != -1:
        st.write("SHIP TO DEBUG:", repr(text[idx:idx+100]))
