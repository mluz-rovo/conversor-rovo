import streamlit as st
import pandas as pd
import io
import pdfplumber
import re

st.set_page_config(page_title="ROVO - Universal Converter", page_icon="🚀", layout="wide")

st.sidebar.title("🚀 ROVO MENU")
client = st.sidebar.selectbox("Select Client", ["Stussy", "Supreme", "Studio Nicholson"])

ref_manual = ""
des_manual = ""

if client == "Supreme":
    st.sidebar.write("---")
    st.sidebar.subheader("📝 Supreme Fixed Data")
    ref_manual = st.sidebar.text_input("Reference (PHC)", placeholder="e.g., FW24-001")
    des_manual = st.sidebar.text_input("Designation (PHC)", placeholder="e.g., Box Logo Hooded")
    st.sidebar.caption("These values will be applied to all rows in the file.")

st.title(f"📦 Converter: {client}")

# ===========================================================================
# STUDIO NICHOLSON — CONSTANTES
# ===========================================================================
SIZE_REFS      = ["XXS", "XS", "S", "M", "L", "XL", "XXL",
                  "UK4", "UK6", "UK8", "UK10", "UK12", "UK14",
                  "UK4/IT36", "UK6/IT38", "UK8/IT40", "UK10/IT42", "UK12/IT44", "UK14/IT46"]
SKIP_LINES     = ["TOTAL QTY", "FIRST/MAKE", "SUB-TOTAL", "TOTAL COST", "QTY COST TOTAL"]
COLOR_JUNK     = {
    "JERSEY", "MICRO", "RIB", "SHORT", "SLEEVE", "NECK", "VEST", "HENLEY",
    "COTTON", "BRANDED", "BOXY", "FIT", "T-SHIRT", "QTY", "COST", "TOTAL",
    "FIRST", "MAKE", "-", "–", "SORIN", "VOTAN", "LAY"
}

def extract_code(text: str) -> str:
    """Extrai o código de referência do modelo, ex: SNW-1868, SNM-1066."""
    m = re.search(r"(S[NW]W?-\d+|SN-\d+)", text, re.IGNORECASE)
    return m.group(1).upper() if m else ""


# ===========================================================================
# PDF DE PREÇOS (com € e Ship To)
# Devolve dict: {(code, color): unit_price}
# ===========================================================================
def parse_prices_pdf(pdf_file) -> dict:
    prices = {}
    current_code  = ""
    current_sizes = []

    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text  = page.extract_text() or ""
            lines = text.split("\n")

            for line in lines:
                l_up = line.upper().strip()
                if not l_up:
                    continue

                # Linha de modelo (começa com LAY, SNW, SNM, SN)
                if re.match(r"(LAY\s+)?(SNW|SNM|SN)-", l_up):
                    current_code = extract_code(line)
                    # Tamanhos presentes nesta linha
                    seen, current_sizes = set(), []
                    for token in line.split():
                        t = token.upper().strip(".,")
                        if t in SIZE_REFS and t not in seen:
                            current_sizes.append(t)
                            seen.add(t)
                    continue

                # Linha de preço (contém €)
                if "€" in line and current_code:
                    price_matches = re.findall(r"€\s*([\d,\.]+)", line)
                    if not price_matches:
                        continue
                    unit_price = float(price_matches[0].replace(",", ""))

                    # Cor: tokens antes dos números excluindo junk
                    before_euro = line.split("€")[0].strip()
                    before_nums = re.split(r"\s+\d", before_euro)[0]
                    color_tokens = [
                        t for t in before_nums.split()
                        if t.upper().strip("–-") not in COLOR_JUNK
                        and not re.match(r"^[\d.,]+$", t)
                        and len(t) > 1
                    ]
                    color = " ".join(color_tokens[-2:]).upper()
                    prices[(current_code, color)] = unit_price

    return prices


# ===========================================================================
# PDF DE QUANTIDADES (com UK sizes e SHIP TO no cabeçalho)
# Devolve lista de dicts com todos os campos excepto unit_price
# ===========================================================================
def parse_quantities_pdf(pdf_file) -> list:
    rows = []
    current_dest  = "See PDF"
    current_code  = ""
    current_model = ""
    current_sizes = []

    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text  = page.extract_text() or ""
            lines = text.split("\n")

            for line in lines:
                l_up = line.upper().strip()
                if not l_up:
                    continue

                # Destino — "SHIP TO ..." no cabeçalho
                if l_up.startswith("SHIP TO"):
                    # Tira "SHIP TO" e fica com o resto, ex: "UK WAREHOUSE - TU PACK"
                    dest_raw = re.sub(r"^SHIP\s+TO\s*", "", line, flags=re.I).strip()
                    if " - " in dest_raw:
                        current_dest = dest_raw.split(" - ")[-1].strip()
                    else:
                        current_dest = dest_raw
                    continue

                # Linha de modelo
                if re.match(r".*(SNW|SNM|SN)\s*-\s*\d+", l_up):
                    current_code  = extract_code(line)
                    current_model = line.strip()
                    # Tamanhos nesta linha (ex: UK4/IT36 UK6/IT38 ...)
                    seen, current_sizes = set(), []
                    for token in line.split():
                        t = token.upper().strip(".,")
                        if t in SIZE_REFS and t not in seen:
                            current_sizes.append(t)
                            seen.add(t)
                    continue

                # Linha de quantidades: tem números e uma cor reconhecível
                # Ignora linhas de totais
                if any(skip in l_up for skip in SKIP_LINES):
                    continue

                if current_code and current_sizes:
                    nums = re.findall(r"\b(\d+)\b", line)
                    if not nums:
                        continue

                    qty_values = [int(n) for n in nums]
                    # Último número = total geral → ignorar
                    qty_values = qty_values[:-1]

                    if not qty_values:
                        continue

                    # Cor: tokens sem números e sem junk
                    color_tokens = [
                        t for t in line.split()
                        if t.upper().strip("–-") not in COLOR_JUNK
                        and not re.match(r"^[\d.,/]+$", t)
                        and len(t) > 1
                        and t.upper() not in SIZE_REFS
                    ]
                    color = " ".join(color_tokens[-2:]).upper()
                    if not color:
                        continue

                    # Alinha qtds pelos últimos N tamanhos (zeros à esquerda omitidos)
                    offset = len(current_sizes) - len(qty_values)
                    for i, size in enumerate(current_sizes):
                        idx = i - offset
                        if idx >= 0 and qty_values[idx] > 0:
                            rows.append({
                                "code":        current_code,
                                "color":       color,
                                "model":       current_model,
                                "size":        size,
                                "qty":         qty_values[idx],
                                "destination": current_dest,
                            })

    return rows


# ===========================================================================
# CRUZAMENTO: quantidades + preços → linhas finais
# ===========================================================================
def merge_sn(qty_rows: list, prices: dict) -> list:
    final = []
    unmatched = set()

    for r in qty_rows:
        key = (r["code"], r["color"])
        unit_price = prices.get(key, 0)
        if unit_price == 0:
            unmatched.add(key)

        final.append({
            "Reference":           "",
            "Designation":         r["model"],
            "Qty":                 r["qty"],
            "Unit Price":          unit_price,
            "Unit Price Currency": 0,
            "VAT Table":           4,
            "Color":               r["color"],
            "Size":                r["size"],
            "TOTAL":               r["qty"] * unit_price,
            "Destination":         r["destination"],
            "CPO No.":             "",
            "SPO No.":             "",
            "Supplier Unit Value": "",
            "Total Supplier":      "",
        })

    if unmatched:
        st.warning(f"⚠️ Preço não encontrado para: {', '.join(str(k) for k in unmatched)}")

    return final


# ===========================================================================
# UPLOADERS
# ===========================================================================
if client == "Studio Nicholson":
    st.info("📎 Faz upload dos dois PDFs: o de **preços** (com €) e o de **quantidades** (com UK sizes).")
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**PDF de Preços** (com €)")
        pdf_prices = st.file_uploader("Upload PDF Preços", type=["pdf"], key="pdf_prices")
    with col2:
        st.markdown("**PDF de Quantidades** (com UK sizes)")
        pdf_qty = st.file_uploader("Upload PDF Quantidades", type=["pdf"], key="pdf_qty")
    uploaded_file = None  # não usado para SN
else:
    uploaded_file = st.file_uploader("Upload file", type=["xlsx"])
    pdf_prices = None
    pdf_qty    = None


# ===========================================================================
# APP PRINCIPAL
# ===========================================================================
cols = [
    "Reference", "Designation", "Qty", "Unit Price",
    "Unit Price Currency", "VAT Table", "Color", "Size",
    "TOTAL", "Destination", "CPO No.", "SPO No.",
    "Supplier Unit Value", "Total Supplier",
]

try:
    data_list = []

    # ── STUSSY ──────────────────────────────────────────────────────────
    if client == "Stussy" and uploaded_file:
        xl = pd.ExcelFile(uploaded_file, engine="openpyxl")
        sheet_name = "Sheet1" if "Sheet1" in xl.sheet_names else xl.sheet_names[0]
        df = xl.parse(sheet_name, header=None)

        for i, row in df.iloc[1:].iterrows():
            if len(row) >= 18:
                q_raw = row[12]
                p_raw = row[17]
                q = pd.to_numeric(q_raw, errors="coerce")
                if isinstance(p_raw, str):
                    p = pd.to_numeric(re.sub(r"[^\d\.]", "", p_raw.replace(",", ".")), errors="coerce")
                else:
                    p = pd.to_numeric(p_raw, errors="coerce")

                if q and q > 0:
                    p_val = p if pd.notna(p) else 0
                    data_list.append({
                        "Reference":           "",
                        "Designation":         row[8] if len(row) > 8 else "",
                        "Qty":                 q,
                        "Unit Price":          0,
                        "Unit Price Currency": p_val,
                        "VAT Table":           4,
                        "Color":               row[7] if len(row) > 7 else "",
                        "Size":                row[9] if len(row) > 9 else "",
                        "TOTAL":               q * p_val,
                        "Destination":         row[4] if len(row) > 4 else "General",
                        "CPO No.":             "",
                        "SPO No.":             "",
                        "Supplier Unit Value": "",
                        "Total Supplier":      "",
                    })

    # ── SUPREME ─────────────────────────────────────────────────────────
    elif client == "Supreme" and uploaded_file:
        xl = pd.ExcelFile(uploaded_file, engine="openpyxl")
        for sheet in xl.sheet_names:
            if "TOTAL" in sheet.upper():
                continue
            df = xl.parse(sheet, header=None)
            sizes = {c: str(df.iloc[14, c]) for c in range(9, 16) if pd.notna(df.iloc[14, c])}

            for start in range(16, len(df), 14):
                dest = str(df.iloc[start, 0]).strip()
                if not dest or dest == "nan":
                    dest = "General"

                for i in range(start + 1, start + 13):
                    if i >= len(df) or pd.isna(df.iloc[i, 6]):
                        continue
                    p = pd.to_numeric(df.iloc[i, 17], errors="coerce")
                    for c_idx, s_name in sizes.items():
                        q = pd.to_numeric(df.iloc[i, c_idx], errors="coerce")
                        if q and q > 0:
                            p_val = p if pd.notna(p) else 0
                            data_list.append({
                                "Reference":           ref_manual,
                                "Designation":         des_manual,
                                "Qty":                 q,
                                "Unit Price":          0,
                                "Unit Price Currency": p_val,
                                "VAT Table":           4,
                                "Color":               df.iloc[i, 6],
                                "Size":                s_name,
                                "TOTAL":               q * p_val,
                                "Destination":         dest,
                                "CPO No.":             "",
                                "SPO No.":             "",
                                "Supplier Unit Value": "",
                                "Total Supplier":      "",
                            })

    # ── STUDIO NICHOLSON ─────────────────────────────────────────────────
    elif client == "Studio Nicholson" and pdf_prices and pdf_qty:
        with st.spinner("A processar PDFs..."):
            prices    = parse_prices_pdf(pdf_prices)
            qty_rows  = parse_quantities_pdf(pdf_qty)
            data_list = merge_sn(qty_rows, prices)

        with st.expander("🔍 Debug: Preços extraídos", expanded=True):
            st.write(f"Total chaves: {len(prices)}")
            st.write(prices)
        with st.expander("🔍 Debug: Quantidades extraídas", expanded=True):
            st.write(f"Total linhas: {len(qty_rows)}")
            st.write(qty_rows[:30])
        with st.expander("🔬 Debug: Linhas cruas PDF Preços", expanded=True):
            with pdfplumber.open(pdf_prices) as pdf:
                for p_num, page in enumerate(pdf.pages):
                    text = page.extract_text() or ""
                    st.markdown(f"**Página {p_num + 1}**")
                    for i, line in enumerate(text.split("\n")):
                        st.code(f"[{i}] {repr(line)}")
        with st.expander("🔬 Debug: Linhas cruas PDF Quantidades", expanded=True):
            with pdfplumber.open(pdf_qty) as pdf:
                for p_num, page in enumerate(pdf.pages):
                    text = page.extract_text() or ""
                    st.markdown(f"**Página {p_num + 1}**")
                    for i, line in enumerate(text.split("\n")):
                        st.code(f"[{i}] {repr(line)}")

    # ── OUTPUT ───────────────────────────────────────────────────────────
    if data_list:
        df_final = pd.DataFrame(data_list).drop_duplicates()

        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as writer:
            for dest in df_final["Destination"].unique():
                safe_name = re.sub(r"[\[\]*:?/\\]", "", str(dest))[:31]
                df_final[df_final["Destination"] == dest][cols].to_excel(
                    writer, index=False, sheet_name=safe_name
                )

        st.success(f"✅ Conversão concluída! {len(data_list)} linhas geradas.")
        st.download_button("⬇️ Download PHC Excel", out.getvalue(), f"IMPORT_{client}.xlsx")

    elif client == "Studio Nicholson" and (not pdf_prices or not pdf_qty):
        pass  # aguarda upload dos dois ficheiros
    else:
        if uploaded_file or (pdf_prices and pdf_qty):
            st.warning("Nenhum dado válido encontrado. Verifica o ficheiro.")

except Exception as e:
    st.error(f"Erro: {e}")
    st.exception(e)
