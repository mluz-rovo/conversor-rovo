import streamlit as st
import pandas as pd
import pdfplumber
import re
import io

# ── Padrões ────────────────────────────────────────────────────────────────────

# Tamanhos UK/IT ex: "UK4 / IT36"
UKSIZE_RE = re.compile(r"UK\d+\s*/\s*IT\d+", re.IGNORECASE)
# Tamanhos standard ex: "XXS", "XS", "S", "M", "L", "XL", "XXL"
STD_SIZES = ["XXS", "XS", "S", "M", "L", "XL", "XXL"]
STD_SIZE_RE = re.compile(r"\b(XXS|XS|S|M|L|XL|XXL)\b")

# Modelo: qualquer nome antes de SNW/SNM/SN seguido de número
MODEL_RE = re.compile(r"([A-Z]+\s+SN[WM]?\s*-\s*\d+)", re.IGNORECASE)

# Preço unitário (primeiro € que aparece antes do Total Cost)
PRICE_RE = re.compile(r"€\s*([\d,\.]+)")

# Bloco de produto — dois sabores conforme tipo de tamanhos
# Versão UK/IT
PRODUCT_UK = re.compile(
    r"([A-Z][\w\s]+-\s*\d+[\w\s\-]*?)"           # modelo
    r"((?:UK\d+\s*/\s*IT\d+\s*)+)"                # tamanhos UK
    r"(JERSEY\s*-\s*[\w\s]+?)\s+"                 # descrição+cor
    r"((?:(?:\d+|-)\s+){2,})"                     # quantidades
    r"\d+\s+"                                      # qty total
    r"€\s*([\d,\.]+)\s+"                          # preço unit
    r"€\s*[\d,\.]+",                               # total (ignorado)
    re.IGNORECASE
)

# Versão STD sizes (XS S M L XL XXL / XXS...)
PRODUCT_STD = re.compile(
    r"([A-Z][\w\s]+-\s*\d+[\w\s\-]*?)"           # modelo
    r"((?:(?:XXS|XS|S|M|L|XL|XXL)\s*)+)"         # tamanhos STD
    r"(JERSEY\s*-\s*[\w\s]+?)\s+"                 # descrição+cor
    r"((?:(?:\d+|-)\s+){2,})"                     # quantidades
    r"\d+\s+"                                      # qty total
    r"€\s*([\d,\.]+)\s+"                          # preço unit
    r"€\s*[\d,\.]+",                               # total (ignorado)
    re.IGNORECASE
)


def extract_ship_to(text):
    """Captura o nome do destino logo após 'Ship To:' até à morada (rua/número)."""
    m = re.search(r"Ship To:\s*(.+?)(?=\s+\w+\s+\d|\s+\d{4}|\s+Bunschotenweg|\s+Rua |\s+Street|\s+Avenue|\s+Road|\s+Docket)", text, re.IGNORECASE)
    if m:
        return m.group(1).strip()
    # Fallback: tudo entre "Ship To:" e "Docket"
    m2 = re.search(r"Ship To:\s*(.+?)Docket", text, re.IGNORECASE | re.DOTALL)
    if m2:
        return m2.group(1).split()[0:4]
    return "Ver PDF"


def parse_modelo(texto):
    """Extrai 'NOME SNW-XXXX' normalizando espaços à volta do traço."""
    m = MODEL_RE.search(texto)
    if m:
        return re.sub(r"\s*-\s*", "-", m.group(1).strip())
    return texto.strip()


def parse_color(descricao):
    """Remove prefixo de descrição — o que sobra é a cor."""
    noise = {"JERSEY", "-", "SHORT", "SLEEVE", "HENLEY", "SCOOP", "NECK",
             "VEST", "BOXY", "FIT", "T-SHIRT", "COTTON", "BRANDED", "CREW",
             "LONG", "L/S", "FLEECE", "SWEATSHIRT", "POLO", "TOUCH", "SOFT"}
    partes = descricao.strip().split()
    cor = [p for p in partes if p.upper() not in noise]
    return " ".join(cor) if cor else descricao.strip()


def process_match(m, size_type, rows, destino):
    modelo    = parse_modelo(m.group(1))
    sizes_raw = m.group(2).strip()
    desc_cor  = m.group(3).strip()
    qtys_raw  = m.group(4).strip().split()
    preco     = float(m.group(5).replace(",", ""))
    cor       = parse_color(desc_cor)

    if size_type == "UK":
        tamanhos = [f"UK{a}/IT{b}" for a, b in
                    re.findall(r"UK(\d+)\s*/\s*IT(\d+)", sizes_raw, re.IGNORECASE)]
    else:
        tamanhos = STD_SIZE_RE.findall(sizes_raw)

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


def extract_studio_nicholson(pdf_file):
    rows = []
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            texto   = page.extract_text() or ""
            destino = extract_ship_to(texto)

            for m in PRODUCT_UK.finditer(texto):
                process_match(m, "UK", rows, destino)

            for m in PRODUCT_STD.finditer(texto):
                process_match(m, "STD", rows, destino)

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
            pass  # Lógica Stussy/Supreme mantida aqui

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

            st.success(f"✅ {len(df)} done!")
            st.dataframe(df[cols], use_container_width=True)
            st.download_button("⬇️ Download Excel", out.getvalue(), "IMPORTAR_PHC.xlsx")
        elif arquivo:
            st.warning("⚠️ Nenhum dado extraído. Verifica o PDF ou o cliente selecionado.")

    except Exception as e:
        st.error(f"Erro: {e}")
        st.exception(e)

elif arquivo.name.endswith(".pdf") and cliente == "Studio Nicholson":
    with pdfplumber.open(arquivo) as pdf:
        for i, page in enumerate(pdf.pages):
            st.subheader(f"Página {i+1} — Texto raw")
            st.text(page.extract_text())
