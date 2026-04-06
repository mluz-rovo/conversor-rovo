import streamlit as st
import pandas as pd
import pdfplumber
import re
import io

# ── Padrões ────────────────────────────────────────────────────────────────────

# Modelo: ex "KUYTO SNM - 1485" ou "THALEN SN - 1869"
MODEL_RE = re.compile(r"([A-Z]+\s+SN[WM]?\s*-\s*\d+)", re.IGNORECASE)

PRICE_RE = re.compile(r"€\s*([\d,\.]+)")

# Bloco UK: modelo ... UK sizes ... "Qty Cost Total Cost" ... descrição+cor ... qtds ... total € unit € total
PRODUCT_UK = re.compile(
    r"([A-Z][A-Z\s]+-\s*\d+[A-Z\s\-]*?)"
    r"((?:UK\d+\s*/\s*IT\d+\s*)+)"
    r"Qty\s+Cost\s+Total\s+Cost\s+"
    r"(JERSEY\s*-\s*[A-Z/\s]+?)\s+"
    r"((?:(?:\d+|-)\s+){2,})"
    r"(\d+)\s+"
    r"€\s*([\d,\.]+)\s+"
    r"€\s*[\d,\.]+",
    re.IGNORECASE
)

# Bloco STD: modelo ... STD sizes ... "Qty Cost Total Cost" ... descrição+cor ... qtds ... total € unit € total
PRODUCT_STD = re.compile(
    r"([A-Z][A-Z\s]+-\s*\d+[A-Z\s\-]*?)"
    r"((?:(?:XXS|XS|S|M|L|XL|XXL)\s+){2,})"
    r"Qty\s+Cost\s+Total\s+Cost\s+"
    r"(JERSEY\s*-\s*[A-Z/\s]+?)\s+"
    r"((?:(?:\d+|-)\s+){2,})"
    r"(\d+)\s+"
    r"€\s*([\d,\.]+)\s+"
    r"€\s*[\d,\.]+",
    re.IGNORECASE
)


def extract_ship_to(text):
    """Captura texto entre 'Ship To:' e 'Docket Number', pega só a primeira parte."""
    m = re.search(r"Ship To:\s*(.+?)Docket Number", text, re.IGNORECASE | re.DOTALL)
    if m:
        # O texto está numa linha contínua: "NL Rotterdam - SEKO Bunschotenweg 160 KC ROTTERDAM..."
        # Queremos só até ao primeiro token que pareça morada (rua com número)
        raw = m.group(1).strip()
        # Corta antes de qualquer palavra seguida de número (ex: "Bunschotenweg 160")
        corte = re.split(r"\b[A-Za-z]+\s+\d+\b", raw)
        return corte[0].strip() if corte else raw
    return "Ver PDF"


def parse_modelo(texto):
    """Extrai 'NOME SN(W/M)-XXXX' e normaliza."""
    m = MODEL_RE.search(texto)
    if m:
        return re.sub(r"\s*-\s*", "-", m.group(1).strip())
    return texto.strip()


def parse_color(descricao):
    """Remove palavras de descrição — o que sobra é a cor."""
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
    preco     = float(m.group(6).replace(",", ""))
    cor       = parse_color(desc_cor)

    if size_type == "UK":
        tamanhos = [f"UK{a}/IT{b}" for a, b in
                    re.findall(r"UK(\d+)\s*/\s*IT(\d+)", sizes_raw, re.IGNORECASE)]
    else:
        tamanhos = re.findall(r"XXS|XS|S|M|L|XL|XXL", sizes_raw, re.IGNORECASE)
        tamanhos = [t.upper() for t in tamanhos]

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

            st.success(f"✅ {len(df)} linhas extraídas com sucesso!")
            st.dataframe(df[cols], use_container_width=True)
            st.download_button("⬇️ Download Excel", out.getvalue(), "IMPORTAR_PHC.xlsx")
        elif arquivo:
            st.warning("⚠️ Nenhum dado extraído. Verifica o PDF ou o cliente selecionado.")

    except Exception as e:
        st.error(f"Erro: {e}")
        st.exception(e)
