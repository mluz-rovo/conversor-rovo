import streamlit as st
import pandas as pd
import io
import pdfplumber
import re

st.set_page_config(page_title="ROVO - Universal Converter", page_icon="🚀", layout="wide")

st.sidebar.title("🚀 ROVO MENU")
client = st.sidebar.selectbox("Select Client", ["Stussy", "Supreme", "Studio Nicholson"])

# ===========================================================================
# SIDEBAR — campos por cliente
# ===========================================================================
ref_manual = ""
des_manual = ""

if client == "Supreme":
    st.sidebar.write("---")
    st.sidebar.subheader("📝 Supreme Fixed Data")
    ref_manual = st.sidebar.text_input("Reference (PHC)", placeholder="e.g., FW24-001")
    des_manual = st.sidebar.text_input("Designation (PHC)", placeholder="e.g., Box Logo Hooded")
    st.sidebar.caption("These values will be applied to all rows in the file.")

stussy_ref_map = {}
stussy_des_map = {}
if client == "Stussy" and st.session_state.get("stussy_models"):
    st.sidebar.write("---")
    st.sidebar.subheader("📝 Stussy — Referências PHC")
    for model in st.session_state["stussy_models"]:
        st.sidebar.caption(model)
        stussy_ref_map[model] = st.sidebar.text_input(
            "Reference (PHC)", key=f"ref_{model}", placeholder="e.g., AW24-001"
        )
        stussy_des_map[model] = st.sidebar.text_input(
            "Designation (PHC)", key=f"des_{model}", placeholder="e.g., Box Logo Tee"
        )

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
MODEL_RE = re.compile(r"(SNW|SNM|SN)\s*[-–]\s*\d+", re.IGNORECASE)

def extract_code(text: str) -> str:
    m = re.search(r"(S[NW]W?\s*[-–]\s*\d+|SN\s*[-–]\s*\d+)", text, re.IGNORECASE)
    return re.sub(r"\s*[-–]\s*", "-", m.group(1)).upper() if m else ""

def parse_quantities_pdf(pdf_file) -> list:
    rows = []
    current_dest  = "See PDF"
    current_code  = ""
    current_model = ""
    current_sizes = []
    log = []

    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text  = page.extract_text() or ""
            lines = text.split("\n")

            for idx, line in enumerate(lines):
                l_up = line.upper().strip()
                if not l_up:
                    continue

                if l_up.startswith("SHIP TO"):
                    dest_raw = re.sub(r"^SHIP\s+TO\s*", "", line, flags=re.I)
                    dest_raw = re.sub(r"Ship\s+To:.*$", "", dest_raw, flags=re.I).strip()
                    if " - " in dest_raw:
                        dest_raw = dest_raw.split(" - ")[-1].strip()
                    if len(dest_raw) < 3 and idx + 1 < len(lines):
                        dest_raw = lines[idx + 1].strip()
                    if dest_raw:
                        current_dest = dest_raw
                    log.append(f"DEST → {current_dest}")
                    continue

                if MODEL_RE.search(line):
                    current_code  = extract_code(line)
                    current_model = re.split(r"\s+Qty\b", line, flags=re.I)[0].strip()
                    current_sizes = []
                    log.append(f"MODEL → {current_code} | {current_model}")
                    continue

                if re.search(r"UK\s*\d+", line, re.IGNORECASE) and re.search(r"IT\s*\d+", line, re.IGNORECASE):
                    raw = re.sub(r"\s+", "", line.upper())
                    current_sizes = re.findall(r"UK\d+/IT\d+", raw)
                    log.append(f"SIZES → {current_sizes}")
                    continue

                if any(skip in l_up for skip in SKIP_LINES):
                    continue

                if not current_code or not current_sizes:
                    continue

                PRODUCT_WORDS = {"JERSEY", "KNIT", "WOVEN", "DENIM", "FLEECE", "TWILL"}
                first_word = l_up.split()[0] if l_up.split() else ""
                if first_word not in PRODUCT_WORDS:
                    continue

                before_euro = line.split("€")[0] if "€" in line else line
                normalized  = re.sub(r"(?<!\w)[-–](?!\w)", " ", before_euro)
                nums = re.findall(r"\b(\d+)\b", normalized)
                qty_values = [int(n) for n in nums][:-1] if len(nums) > 1 else [int(n) for n in nums]

                before_nums = re.split(r"\s+\d", normalized)[0]
                color_tokens = [
                    t for t in before_nums.split()
                    if t.upper().strip("–-") not in COLOR_JUNK
                    and not re.match(r"^[\d.,/]+$", t)
                    and len(t) > 1
                ]
                color = " ".join(color_tokens[-2:]).upper() if color_tokens else ""

                if not color or not qty_values:
                    continue

                offset = len(current_sizes) - len(qty_values)
                for i, size in enumerate(current_sizes):
                    i2 = i - offset
                    if i2 >= 0 and qty_values[i2] > 0:
                        rows.append({
                            "code":        current_code,
                            "model":       current_model,
                            "color":       color,
                            "size":        size,
                            "qty":         qty_values[i2],
                            "destination": current_dest,
                        })

    with st.expander("🐛 Debug linha a linha", expanded=False):
        for entry in log:
            st.text(entry)

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

def make_excel(df_final, group_col):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        for val in df_final[group_col].unique():
            safe = re.sub(r"[\[\]*:?/\\]", "", str(val))[:31]
            df_final[df_final[group_col] == val][cols].to_excel(
                writer, index=False, sheet_name=safe
            )
    return out.getvalue()

# ===========================================================================
# STUSSY
# ===========================================================================
if client == "Stussy":
    uploaded_file = st.file_uploader("Upload file", type=["xlsx"])

    if uploaded_file:
        if st.button("🔍 Analisar Ficheiro", type="secondary"):
            xl = pd.ExcelFile(uploaded_file, engine="openpyxl")
            sheet_name = "Sheet1" if "Sheet1" in xl.sheet_names else xl.sheet_names[0]
            df = xl.parse(sheet_name, header=None)
            models_found = df.iloc[1:][8].dropna().astype(str).str.strip().unique().tolist()
            models_found = [m for m in models_found if m and m != "nan"]
            st.session_state["stussy_models"] = models_found
            st.session_state["stussy_df"]     = df
            st.info(f"✅ {len(models_found)} modelo(s) encontrado(s). Preenche as referências no sidebar.")

    if st.session_state.get("stussy_df") is not None:
        if st.button("✅ Gerar Excel"):
            try:
                df = st.session_state["stussy_df"]
                data_list = []
                for i, row in df.iloc[1:].iterrows():
                    if len(row) >= 18:
                        q = pd.to_numeric(row[12], errors="coerce")
                        p_raw = row[17]
                        if isinstance(p_raw, str):
                            p = pd.to_numeric(re.sub(r"[^\d\.]", "", p_raw.replace(",", ".")), errors="coerce")
                        else:
                            p = pd.to_numeric(p_raw, errors="coerce")
                        if q and q > 0:
                            p_val  = p if pd.notna(p) else 0
                            po_raw = row[2] if len(row) > 2 else ""
                            po     = str(po_raw).strip() if pd.notna(po_raw) else "General"
                            model  = str(row[8]).strip() if len(row) > 8 else ""
                            data_list.append({
                                "Reference":           stussy_ref_map.get(model, ""),
                                "Designation":         stussy_des_map.get(model, model),
                                "Qty":                 q,
                                "Unit Price":          0,
                                "Unit Price Currency": p_val,
                                "VAT Table":           4,
                                "Color":               row[7] if len(row) > 7 else "",
                                "Size":                row[9] if len(row) > 9 else "",
                                "TOTAL":               q * p_val,
                                "Destination":         row[4] if len(row) > 4 else "General",
                                "PO":                  po,
                                "CPO No.":             "",
                                "SPO No.":             "",
                                "Supplier Unit Value": "",
                                "Total Supplier":      "",
                            })

                df_final = pd.DataFrame(data_list).drop_duplicates()
                excel    = make_excel(df_final, "PO")
                st.session_state["stussy_excel"] = excel
                st.success(f"✅ Conversão concluída! {len(data_list)} linhas geradas.")
            except Exception as e:
                st.error(f"Erro: {e}")
                st.exception(e)

        if st.session_state.get("stussy_excel"):
            st.download_button(
                "⬇️ Download PHC Excel",
                st.session_state["stussy_excel"],
                "IMPORT_Stussy.xlsx"
            )

# ===========================================================================
# SUPREME
# ===========================================================================
elif client == "Supreme":
    uploaded_file = st.file_uploader("Upload file", type=["xlsx"])

    if uploaded_file:
        try:
            data_list = []
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
                excel    = make_excel(df_final, "Destination")
                st.success(f"✅ Conversão concluída! {len(data_list)} linhas geradas.")
                st.download_button("⬇️ Download PHC Excel", excel, "IMPORT_Supreme.xlsx")
            else:
                st.warning("Nenhum dado válido encontrado.")
        except Exception as e:
            st.error(f"Erro: {e}")
            st.exception(e)

# ===========================================================================
# STUDIO NICHOLSON
# ===========================================================================
elif client == "Studio Nicholson":
    uploaded_file = st.file_uploader("Upload PDF Quantidades", type=["pdf"])

    if uploaded_file:
        file_bytes = uploaded_file.read()
        qty_rows   = parse_quantities_pdf(io.BytesIO(file_bytes))
        if qty_rows:
            st.session_state["sn_rows"] = qty_rows
        else:
            st.warning("Nenhum dado encontrado no PDF.")

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
                excel    = make_excel(df_final, "Destination")
                st.success(f"✅ Conversão concluída! {len(data_list)} linhas geradas.")
                st.session_state["sn_excel"] = excel
            except Exception as e:
                st.error(f"Erro: {e}")
                st.exception(e)

        if st.session_state.get("sn_excel"):
            st.download_button(
                "⬇️ Download PHC Excel",
                st.session_state["sn_excel"],
                "IMPORT_StudioNicholson.xlsx"
            )
