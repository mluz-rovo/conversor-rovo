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

file_format = ["xlsx", "pdf"] if client == "Studio Nicholson" else ["xlsx"]
uploaded_file = st.file_uploader("Upload file", type=file_format)


# ===========================================================================
# STUDIO NICHOLSON — CONSTANTES
# ===========================================================================
SIZE_REFS      = ["XXS", "XS", "S", "M", "L", "XL", "XXL",
                  "UK4", "UK6", "UK8", "UK10", "UK12", "UK14"]
SKIP_LINES     = ["TOTAL QTY", "FIRST/MAKE", "SUB-TOTAL", "TOTAL COST", "QTY COST TOTAL"]
MODEL_PREFIXES = ["SNW -", "SNM -", "SN -", "LAY "]
COLOR_JUNK     = {
    "JERSEY", "MICRO", "RIB", "SHORT", "SLEEVE", "NECK", "VEST", "HENLEY",
    "COTTON", "BRANDED", "BOXY", "FIT", "T-SHIRT", "QTY", "COST", "TOTAL",
    "FIRST", "MAKE", "-", "–"
}


# ---------------------------------------------------------------------------
# PASS 1 — Extração fiel do PDF em eventos
# ---------------------------------------------------------------------------
def sn_extract_raw(pdf_file) -> list:
    events = []

    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text  = page.extract_text() or ""
            lines = text.split("\n")

            # Destino (Ship To)
            if "Ship To:" in text:
                after      = text.split("Ship To:", 1)[1]
                dest_lines = [l.strip() for l in after.split("\n") if l.strip()]
                skip_dest  = re.compile(
                    r"docket|payment|terms|order|season|ref|po\s*#|date",
                    re.IGNORECASE
                )
                dest_line = next(
                    (l for l in dest_lines if not skip_dest.search(l)), None
                )
                if dest_line:
                    if " - " in dest_line:
                        dest_line = dest_line.split(" - ")[-1].strip()
                    events.append({"type": "destination", "dest": dest_line})

            # Percorre linhas
            current_sizes = []  # lista de strings, ex: ["XS","S","M","L","XL","XXL"]
            for line in lines:
                l_up = line.upper().strip()
                if not l_up:
                    continue

                # Linha de modelo — tem prioridade sobre SKIP_LINES
                is_model = any(l_up.startswith(pfx.upper()) for pfx in MODEL_PREFIXES)

                # Ignora linhas de totais/cabeçalhos — mas só se não for linha de modelo
                if not is_model and any(skip in l_up for skip in SKIP_LINES):
                    continue

                if is_model:
                    size_pattern = r"\b(" + "|".join(SIZE_REFS) + r"|Qty|Cost|Total|First)\b"
                    model = re.split(size_pattern, line, flags=re.I)[0].strip()
                    events.append({"type": "model", "model": model})
                    # Tamanhos pela ordem em que aparecem na linha
                    # Tokeniza a linha e filtra só os tokens que são tamanhos válidos
                    # Usa a ordem do texto, não do SIZE_REFS, para preservar sequência real
                    seen = set()
                    current_sizes = []
                    for token in line.split():
                        t = token.upper().strip(".,")
                        if t in SIZE_REFS and t not in seen:
                            current_sizes.append(t)
                            seen.add(t)
                    continue

                # Linha de preço/cor (contém €)
                if "€" in line:
                    ev = _parse_price_row(line, current_sizes)
                    if ev:
                        events.append(ev)

    return events


def _parse_price_row(line: str, sizes: list) -> dict | None:
    """
    Formato: 'JERSEY - BRANDED BOXY FIT T-SHIRT BLACK – 6 11 10 7 1 35 € 13.95 € 488.25'
    sizes: lista ordenada de strings, ex: ["XS","S","M","L","XL","XXL"]
    """
    if "€" not in line or not sizes:
        return None

    # Preço unitário — primeiro valor após €
    price_matches = re.findall(r"€\s*([\d,\.]+)", line)
    if not price_matches:
        return None
    unit_price = float(price_matches[0].replace(",", ""))

    # Parte antes do primeiro €
    before_euro = line.split("€")[0].strip()

    # Quantidades: todos os inteiros antes do €
    nums = re.findall(r"\b(\d+)\b", before_euro)
    if not nums:
        return None
    qty_values = [int(n) for n in nums]
    # O último número é o total geral — ignorar
    qty_values = qty_values[:-1]

    # Associar pela ordem dos tamanhos
    # Se há menos qtds que tamanhos, alinha pela direita (zeros à esquerda omitidos)
    size_quantities = {}
    offset = len(sizes) - len(qty_values)
    for i, size in enumerate(sizes):
        idx = i - offset
        if idx >= 0 and qty_values[idx] > 0:
            size_quantities[size] = qty_values[idx]

    if not size_quantities:
        return None

    # Cor: tokens antes dos números, excluindo junk
    before_nums = re.split(r"\s+\d", before_euro)[0]
    color_tokens = [
        t for t in before_nums.split()
        if t.upper().strip("–-") not in COLOR_JUNK
        and not re.match(r"^[\d.,]+$", t)
        and len(t) > 1
    ]
    color = " ".join(color_tokens[-2:])

    return {
        "type": "price_row",
        "color": color,
        "unit_price": unit_price,
        "size_quantities": size_quantities,
    }


# ---------------------------------------------------------------------------
# PASS 2 — Transforma eventos em linhas finais
# ---------------------------------------------------------------------------
def sn_transform(events: list) -> list:
    rows = []
    current_dest  = "See PDF"
    current_model = ""

    for ev in events:
        if ev["type"] == "destination":
            current_dest = ev["dest"]
        elif ev["type"] == "model":
            current_model = ev["model"]
        elif ev["type"] == "price_row":
            for size, qty in ev["size_quantities"].items():
                rows.append({
                    "Reference":           "",
                    "Designation":         current_model,
                    "Qty":                 qty,
                    "Unit Price":          ev["unit_price"],
                    "Unit Price Currency": 0,
                    "VAT Table":           4,
                    "Color":               ev["color"],
                    "Size":                size,
                    "TOTAL":               qty * ev["unit_price"],
                    "Destination":         current_dest,
                    "CPO No.":             "",
                    "SPO No.":             "",
                    "Supplier Unit Value": "",
                    "Total Supplier":      "",
                })

    return rows


# ---------------------------------------------------------------------------
# DEBUG
# ---------------------------------------------------------------------------
def show_debug(events: list):
    with st.expander("🔍 Debug: Eventos extraídos (Pass 1)", expanded=False):
        for i, ev in enumerate(events):
            st.write(f"**{i}** — `{ev['type']}`", ev)

def show_raw_lines(pdf_file):
    with st.expander("🔬 Debug: Linhas cruas do PDF", expanded=True):
        with pdfplumber.open(pdf_file) as pdf:
            for p_num, page in enumerate(pdf.pages):
                text = page.extract_text() or ""
                st.markdown(f"**Página {p_num + 1}**")
                for i, line in enumerate(text.split("\n")):
                    has_price = "€" in line or "EUR" in line
                    if has_price or any(pfx.upper() in line.upper() for pfx in MODEL_PREFIXES):
                        st.code(f"[{i}] {repr(line)}")


# ===========================================================================
# APP PRINCIPAL
# ===========================================================================
if uploaded_file:
    try:
        data_list = []

        # ── STUSSY ──────────────────────────────────────────────────────────
        if uploaded_file.name.endswith(".xlsx") and client == "Stussy":
            xl = pd.ExcelFile(uploaded_file, engine="openpyxl")
            sheet_name = "Sheet1" if "Sheet1" in xl.sheet_names else xl.sheet_names[0]
            df = xl.parse(sheet_name, header=None)

            for i, row in df.iloc[1:].iterrows():
                if len(row) >= 18:
                    q_raw = row[12]   # Coluna M
                    p_raw = row[17]   # Coluna R

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
        elif uploaded_file.name.endswith(".xlsx") and client == "Supreme":
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
        elif uploaded_file.name.endswith(".pdf") and client == "Studio Nicholson":
            events    = sn_extract_raw(uploaded_file)
            show_debug(events)
            show_raw_lines(uploaded_file)
            data_list = sn_transform(events)

        # ── OUTPUT ───────────────────────────────────────────────────────────
        cols = [
            "Reference", "Designation", "Qty", "Unit Price",
            "Unit Price Currency", "VAT Table", "Color", "Size",
            "TOTAL", "Destination", "CPO No.", "SPO No.",
            "Supplier Unit Value", "Total Supplier",
        ]

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
        else:
            st.warning("Nenhum dado válido encontrado. Verifica o ficheiro.")

    except Exception as e:
        st.error(f"Erro: {e}")
        st.exception(e)
