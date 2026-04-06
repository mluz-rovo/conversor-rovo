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
SKIP_LINES = ["TOTAL QTY", "FIRST/MAKE", "SUB-TOTAL", "TOTAL COST", "QTY COST TOTAL"]
COLOR_JUNK = {
    "JERSEY", "MICRO", "RIB", "SHORT", "SLEEVE", "NECK", "VEST", "HENLEY",
    "COTTON", "BRANDED", "BOXY", "FIT", "T-SHIRT", "QTY", "COST", "TOTAL",
    "FIRST", "MAKE", "-", "–", "SORIN", "VOTAN", "LAY", "SCOOP",
    "PRODUCTION", "MADE", "LOCATION", "UNITED", "KINGDOM", "KOREA", "SOUTH",
    "PRINTED", "REG", "OFFICE", "VAT", "PAGE"
}
MODEL_RE      = re.compile(r"(SNW|SNM|SN)\s*[-–]\s*\d+", re.IGNORECASE)
PRODUCT_WORDS = {"JERSEY", "KNIT", "WOVEN", "DENIM", "FLEECE", "TWILL"}
SIZE_REFS_STD = ["XXS", "XS", "S", "M", "L", "XL", "XXL"]
SIZE_IT_KEYS  = ["IT36", "IT38", "IT40", "IT42", "IT44", "IT46"]


# ===========================================================================
# STUDIO NICHOLSON — FUNÇÕES
# ===========================================================================
def extract_code(text):
    m = re.search(r"(S[NW]W?\s*[-–]\s*\d+|SN\s*[-–]\s*\d+)", text, re.IGNORECASE)
    return re.sub(r"\s*[-–]\s*", "-", m.group(1)).upper() if m else ""


def parse_size_line_std(line):
    seen, sizes = set(), []
    for token in line.split():
        t = token.upper().strip(".,")
        if t in SIZE_REFS_STD and t not in seen:
            sizes.append(t)
            seen.add(t)
    return sizes


def parse_quantities_pdf(pdf_source):
    rows = []
    current_dest  = "See PDF"
    current_code  = ""
    current_model = ""
    current_sizes = []

    with pdfplumber.open(pdf_source) as pdf:
        for page in pdf.pages:
            text  = page.extract_text() or ""
            lines = text.split("\n")

            for idx, line in enumerate(lines):
                l_up = line.upper().strip()
                if not l_up:
                    continue

                # 1. DESTINO via SHIP TO — usa sempre a linha seguinte
                if l_up.startswith("SHIP TO"):
                    if idx + 1 < len(lines):
                        current_dest = lines[idx + 1].strip()
                    continue

                # 2. DESTINO via "ROVO - ARAUJO IRMAOS <local>" (fallback)
                if l_up.startswith("ROVO -") or l_up.startswith("ROVO–"):
                    if not current_dest or current_dest == "See PDF":
                        parts = line.strip().split()
                        try:
                            irmaos_idx = [p.upper() for p in parts].index("IRMAOS")
                            current_dest = " ".join(parts[irmaos_idx + 1:]).strip()
                        except ValueError:
                            current_dest = " ".join(parts[-2:]).strip()
                    continue

                # 3. MODELO
                if MODEL_RE.search(line):
                    current_code  = extract_code(line)
                    current_model = re.split(r"\s+Qty\b", line, flags=re.I)[0].strip()
                    current_sizes = []
                    continue

                # 4. TAMANHOS — UK/IT ou standard (XS, S, M, L...)
                is_uk  = any(k in l_up for k in SIZE_IT_KEYS)
                is_std = any(f" {s} " in f" {l_up} " for s in SIZE_REFS_STD)

                if is_uk:
                    raw = re.sub(r"\s+", "", line.upper())
                    current_sizes = re.findall(r"UK\d+/IT\d+", raw)
                    st.write(f"SIZES UK: {current_sizes} ← {repr(line[:60])}")
                    continue
                elif is_std and not MODEL_RE.search(line):
                    current_sizes = parse_size_line_std(line)
                    st.write(f"SIZES STD: {current_sizes} ← {repr(line[:60])}")
                    continue

                # 5. SKIP totais
                if any(skip in l_up for skip in SKIP_LINES):
                    st.write(f"SKIP: {repr(line[:60])}")
                    continue

                # 6. QUANTIDADES
                if not current_code or not current_sizes:
                    st.write(f"NO STATE (code={current_code!r}, sizes={current_sizes}): {repr(line[:60])}")
                    continue

                first_word = l_up.split()[0] if l_up.split() else ""
                if first_word not in PRODUCT_WORDS:
                    st.write(f"NOT PRODUCT ({first_word!r}): {repr(line[:60])}")
                    continue

                st.write(f"✅ QTY LINE: {repr(line[:60])}")

                before_euro = line.split("€")[0] if "€" in line else line
                normalized  = re.sub(r"(?<!\w)[-–](?!\w)", "0", before_euro)
                nums = re.findall(r"\b(\d+)\b", normalized)
                if not nums:
                    continue
                qty_values = [int(n) for n in nums]
                qty_values = qty_values[:-1]                    # remove total
                qty_values = qty_values[:len(current_sizes)]    # trunca ao nº de tamanhos
                if not qty_values:
                    continue

                before_nums = re.split(r"\s+\d", normalized)[0]
                color_tokens = [
                    t for t in before_nums.split()
                    if t.upper().strip("–-") not in COLOR_JUNK
                    and not re.match(r"^[\d.,/]+$", t)
                    and len(t) > 1
                ]
                color = " ".join(color_tokens[-2:]).upper() if color_tokens else ""
                if not color:
                    continue

                for i, size in enumerate(current_sizes):
                    if i < len(qty_values) and qty_values[i] > 0:
                        rows.append({
                            "code":        current_code,
                            "model":       current_model,
                            "color":       color,
                            "size":        size,
                            "qty":         qty_values[i],
                            "destination": current_dest,
                        })

    return rows


# ===========================================================================
# COLUNAS FINAIS
# ===========================================================================
cols = [
    "Reference", "Designation", "Qty", "Unit Price",
    "Unit Price Currency", "VAT Table", "Color", "Size",
    "TOTAL", "Destination", "CPO No.", "SPO No.",
    "Supplier Unit Value", "Total Supplier",
]


# ===========================================================================
# STUSSY & SUPREME
# ===========================================================================
if client in ["Stussy", "Supreme"]:
    uploaded_file = st.file_uploader("Upload file", type=["xlsx"])

    if uploaded_file:
        try:
            data_list = []

            if client == "Stussy":
                xl = pd.ExcelFile(uploaded_file, engine="openpyxl")
                sheet_name = "Sheet1" if "Sheet1" in xl.sheet_names else xl.sheet_names[0]
                df = xl.parse(sheet_name, header=None)
                for i, row in df.iloc[1:].iterrows():
                    if len(row) >= 18:
                        q = pd.to_numeric(row[12], errors="coerce")
                        p_raw = row[17]
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

            elif client == "Supreme":
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


# ===========================================================================
# STUDIO NICHOLSON
# ===========================================================================
elif client == "Studio Nicholson":
    uploaded_file = st.file_uploader("Upload PDF Quantidades", type=["pdf"])

    if uploaded_file:
        try:
            file_bytes = uploaded_file.read()
            qty_rows   = parse_quantities_pdf(io.BytesIO(file_bytes))
            if qty_rows:
                st.session_state["sn_rows"] = qty_rows
            elif "sn_rows" not in st.session_state:
                st.warning("Nenhum dado encontrado no PDF.")
        except Exception as e:
            st.error(f"Erro: {e}")
            st.exception(e)

    if st.session_state.get("sn_rows"):
        qty_rows = st.session_state["sn_rows"]
        models   = sorted({(r["code"], r["model"]) for r in qty_rows}, key=lambda x: x[0])

        st.subheader("💶 Introduz o preço unitário por modelo")
        price_map = {}
        cols_ui   = st.columns(min(len(models), 3))
        for i, (code, model_name) in enumerate(models):
            with cols_ui[i % 3]:
                price = st.number_input(
                    f"{code}",
                    min_value=0.0,
                    step=0.01,
                    format="%.2f",
                    key=f"price_{code}",
                    help=model_name,
                )
                price_map[code] = price

        if st.button("✅ Gerar Excel", type="primary"):
            try:
                data_list = []
                for r in qty_rows:
                    unit_price = price_map.get(r["code"], 0)
                    data_list.append({
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

                df_final = pd.DataFrame(data_list).drop_duplicates()
                out = io.BytesIO()
                with pd.ExcelWriter(out, engine="openpyxl") as writer:
                    for dest in df_final["Destination"].unique():
                        safe_name = re.sub(r"[\[\]*:?/\\]", "", str(dest))[:31]
                        df_final[df_final["Destination"] == dest][cols].to_excel(
                            writer, index=False, sheet_name=safe_name
                        )
                st.success(f"✅ Conversão concluída! {len(data_list)} linhas geradas.")
                st.session_state["sn_excel"] = out.getvalue()
            except Exception as e:
                st.error(f"Erro ao gerar Excel: {e}")
                st.exception(e)

        if st.session_state.get("sn_excel"):
            st.download_button(
                "⬇️ Download PHC Excel",
                st.session_state["sn_excel"],
                "IMPORT_StudioNicholson.xlsx"
            )
