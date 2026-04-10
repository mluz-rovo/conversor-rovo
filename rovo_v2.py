import streamlit as st
import pandas as pd
import io
import pdfplumber
import re

st.set_page_config(page_title="ROVO - Universal Converter", page_icon="🚀", layout="wide")

st.sidebar.title("🚀 ROVO MENU")
client = st.sidebar.selectbox("Select Client", ["Stussy", "Supreme", "Studio Nicholson", "Index"], key="client_select")

# ===========================================================================
# SIDEBAR
# ===========================================================================
ref_manual = ""
des_manual = ""

if client == "Supreme":
    st.sidebar.write("---")
    st.sidebar.subheader("📝 Supreme Fixed Data")
    ref_manual = st.sidebar.text_input("Reference (PHC)", placeholder="e.g., FW24-001")
    des_manual = st.sidebar.text_input("Designation (PHC)", placeholder="e.g., Box Logo Hooded")
    st.sidebar.caption("These values will be applied to all rows in the file.")

if client == "Stussy":
    st.sidebar.write("---")
    st.sidebar.subheader("📝 Stussy — PHC References")
    if st.session_state.get("stussy_models"):
        for model in st.session_state["stussy_models"]:
            st.sidebar.caption(model)
            st.sidebar.text_input("Reference (PHC)", key=f"ref_{model}", placeholder="e.g., AW24-001")
            st.sidebar.text_input("Designation (PHC)", key=f"des_{model}", placeholder="e.g., Box Logo Tee")
    else:
        st.sidebar.caption("Upload a file and click Analyse File.")

st.title(f"📦 Converter: {client}")

# ===========================================================================
# STUDIO NICHOLSON — CONSTANTS
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

    with st.expander("🐛 Debug", expanded=False):
        for entry in log:
            st.text(entry)

    return rows

# ===========================================================================
# COLUNAS FINAIS (PT)
# ===========================================================================
cols = [
    "Referência", "Designação", "Quant.", "Pr.Unit.",
    "Pr.Unit. Moeda", "Tabela de IVA", "Cor", "Tamanho",
    "TOTAL", "Destino", "Nº CPO", "Nº SPO",
    "Valor Unit. Fornecedor", "Total Fornecedor",
    "Data Envio Cliente", "Data Envio Fornecedor", "Notas",
]

def make_row(ref="", des="", qty=0, price=0.0, vat=4, color="", size="",
             dest="", cpo="", spo="", supp_val="", supp_total="",
             client_date="", supp_date="", currency=0, notas=""):
    total = qty * price
    return {
        "Referência":             ref,
        "Designação":             des,
        "Quant.":                 qty,
        "Pr.Unit.":               price,
        "Pr.Unit. Moeda":         currency,
        "Tabela de IVA":          vat,
        "Cor":                    color,
        "Tamanho":                size,
        "TOTAL":                  total,
        "Destino":                dest,
        "Nº CPO":                 cpo,
        "Nº SPO":                 spo,
        "Valor Unit. Fornecedor": supp_val,
        "Total Fornecedor":       supp_total,
        "Data Envio Cliente":     client_date,
        "Data Envio Fornecedor":  supp_date,
        "Notas":                  notas,
    }

def make_excel(df_final, group_col):
    if df_final.empty:
        raise ValueError("No valid data to generate Excel. Please check the file.")
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        for val in df_final[group_col].unique():
            safe = re.sub(r"[\[\]*:?/\\]", "", str(val))[:31]
            output_cols = [c for c in cols if c in df_final.columns]
            df_final[df_final[group_col] == val][output_cols].to_excel(
                writer, index=False, sheet_name=safe if safe else "Sheet1"
            )
    return out.getvalue()

# ===========================================================================
# STUSSY
# ===========================================================================
if client == "Stussy":
    uploaded_file = st.file_uploader("Upload file", type=["xlsx"])

    if uploaded_file:
        if st.session_state.get("stussy_filename") != uploaded_file.name:
            for key in ["stussy_models", "stussy_df"]:
                st.session_state.pop(key, None)
            st.session_state["stussy_filename"] = uploaded_file.name

        if st.button("🔍 Analyse File", type="secondary"):
            xl = pd.ExcelFile(uploaded_file, engine="openpyxl")
            sheet_name = "Sheet1" if "Sheet1" in xl.sheet_names else xl.sheet_names[0]
            df = xl.parse(sheet_name, header=None)
            models_found = df.iloc[1:][8].dropna().astype(str).str.strip().unique().tolist()
            models_found = [m for m in models_found if m and m != "nan"]
            st.session_state["stussy_models"] = models_found
            st.session_state["stussy_df"]     = df
            st.rerun()

    if st.session_state.get("stussy_df") is not None and st.session_state.get("stussy_filename") == (uploaded_file.name if uploaded_file else ""):
        try:
            df = st.session_state["stussy_df"]
            data_list = []
            for i, row in df.iloc[1:].iterrows():
                n = len(row)
                q = pd.to_numeric(row[12], errors="coerce") if n > 12 else None
                p_raw = row[17] if n > 17 else (row[13] if n > 13 else None)
                if isinstance(p_raw, str):
                    p = pd.to_numeric(re.sub(r"[^\d\.]", "", p_raw.replace(",", ".")), errors="coerce")
                else:
                    p = pd.to_numeric(p_raw, errors="coerce") if p_raw is not None else None
                if q and q > 0:
                    p_val  = p if p is not None and pd.notna(p) else 0
                    po_raw = row[2] if n > 2 else ""
                    po     = str(po_raw).strip() if pd.notna(po_raw) else "General"
                    if not po or po == "nan":
                        po = "General"
                    model = str(row[8]).strip() if n > 8 else ""
                    r = make_row(
                        ref      = st.session_state.get(f"ref_{model}", ""),
                        des      = st.session_state.get(f"des_{model}", model),
                        qty      = q,
                        price    = 0,
                        currency = p_val,
                        color    = row[7] if n > 7 else "",
                        size     = row[9] if n > 9 else "",
                        dest     = row[4] if n > 4 else "General",
                    )
                    r["PO"] = po
                    data_list.append(r)

            df_final = pd.DataFrame(data_list).drop_duplicates()
            excel    = make_excel(df_final, "PO")
            st.download_button(
                f"⬇️ Download PHC Excel ({len(data_list)} rows)",
                excel,
                "IMPORT_Stussy.xlsx",
                key="dl_stussy"
            )
        except Exception as e:
            st.error(f"Error: {e}")
            st.exception(e)

# ===========================================================================
# SUPREME
# ===========================================================================
elif client == "Supreme":
    supreme_type  = st.radio("File Type", ["Bulk", "SMS", "TOP"], horizontal=True)
    uploaded_file = st.file_uploader("Upload file", type=["xlsx"])

    if uploaded_file:
        try:
            data_list = []
            xl = pd.ExcelFile(uploaded_file, engine="openpyxl")

            if supreme_type == "Bulk":
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
                                    data_list.append(make_row(
                                        ref=ref_manual, des=des_manual,
                                        qty=q, currency=p_val,
                                        color=df.iloc[i, 6], size=s_name, dest=dest,
                                    ))

            else:  # SMS and TOP
                df   = xl.parse(xl.sheet_names[0], header=None)
                dest = str(df.iloc[3, 7]).strip() if pd.notna(df.iloc[3, 7]) else "General"
                sizes = {c: str(df.iloc[14, c]) for c in range(8, 13)
                         if pd.notna(df.iloc[14, c]) and str(df.iloc[14, c]).strip()}
                current_model = ""
                for i in range(17, len(df)):
                    row = df.iloc[i]
                    if pd.notna(row[0]) and str(row[0]).strip():
                        current_model = str(row[0]).strip()
                    color = str(row[6]).strip() if pd.notna(row[6]) and str(row[6]).strip() else ""
                    if not color or color == "nan":
                        continue
                    p     = pd.to_numeric(row[14], errors="coerce") if pd.notna(row[14]) else 0
                    p_val = p if pd.notna(p) else 0
                    for c_idx, s_name in sizes.items():
                        q = pd.to_numeric(row[c_idx], errors="coerce")
                        if q and q > 0:
                            data_list.append(make_row(
                                ref=ref_manual, des=des_manual,
                                qty=q, currency=p_val,
                                color=color, size=s_name, dest=dest,
                                notas="preço unitário com surcharge" if supreme_type == "SMS" else "",
                            ))

            if data_list:
                df_final = pd.DataFrame(data_list).drop_duplicates()
                excel    = make_excel(df_final, "Destino")
                st.success(f"✅ Conversion complete! {len(data_list)} rows generated.")
                st.download_button("⬇️ Download PHC Excel", excel, f"IMPORT_Supreme_{supreme_type}.xlsx")
            else:
                st.warning("No valid data found. Please check the file.")
        except Exception as e:
            st.error(f"Error: {e}")
            st.exception(e)

# ===========================================================================
# STUDIO NICHOLSON
# ===========================================================================
elif client == "Studio Nicholson":
    uploaded_file = st.file_uploader("Upload PDF", type=["pdf"])

    if uploaded_file:
        file_bytes = uploaded_file.read()
        qty_rows   = parse_quantities_pdf(io.BytesIO(file_bytes))
        if qty_rows:
            st.session_state["sn_rows"] = qty_rows
        else:
            st.warning("No data found in the PDF.")

    if st.session_state.get("sn_rows"):
        qty_rows = st.session_state["sn_rows"]
        models   = sorted({(r["code"], r["model"]) for r in qty_rows}, key=lambda x: x[0])

        st.subheader("💶 Enter unit price per model")
        price_map = {}
        cols_ui   = st.columns(min(len(models), 3))
        for i, (code, model_name) in enumerate(models):
            with cols_ui[i % 3]:
                price = st.number_input(
                    f"{code}", min_value=0.0, step=0.01, format="%.2f",
                    key=f"price_{code}", help=model_name,
                )
                price_map[code] = price

        if st.button("✅ Generate Excel"):
            try:
                data_list = []
                for r in qty_rows:
                    unit_price = price_map.get(r["code"], 0)
                    data_list.append(make_row(
                        des=r["model"], qty=r["qty"], price=unit_price,
                        color=r["color"], size=r["size"], dest=r["destination"],
                    ))
                df_final = pd.DataFrame(data_list).drop_duplicates()
                excel    = make_excel(df_final, "Destino")
                st.success(f"✅ Conversion complete! {len(data_list)} rows generated.")
                st.session_state["sn_excel"] = excel
            except Exception as e:
                st.error(f"Error: {e}")
                st.exception(e)

        if st.session_state.get("sn_excel"):
            st.download_button("⬇️ Download PHC Excel", st.session_state["sn_excel"], "IMPORT_StudioNicholson.xlsx")

# ===========================================================================
# INDEX
# ===========================================================================
elif client == "Index":
    st.info("Fill in the table below and click Download to generate the Excel file.")

    empty_row = {
        "Referência":    "",
        "Designação":    "",
        "Quant.":        0,
        "Pr.Unit.":      0.0,
        "Tabela de IVA": 23,
        "Cor":           "",
        "Tamanho":       "",
        "Delivery Date": "",
        "Nº SPO":        "",
        "Supplier":      "",
    }

    if "index_df" not in st.session_state:
        st.session_state["index_df"] = pd.DataFrame([empty_row] * 5)

    col1, col2 = st.columns([1, 5])
    with col1:
        if st.button("➕ Add Row"):
            st.session_state["index_df"] = pd.concat(
                [st.session_state["index_df"], pd.DataFrame([empty_row])],
                ignore_index=True
            )
        if st.button("🗑️ Clear All"):
            st.session_state["index_df"] = pd.DataFrame([empty_row] * 5)

    edited_df = st.data_editor(
        st.session_state["index_df"],
        use_container_width=True,
        num_rows="dynamic",
        column_config={
            "Referência":    st.column_config.TextColumn("Referência"),
            "Designação":    st.column_config.TextColumn("Designação"),
            "Quant.":        st.column_config.NumberColumn("Quant.", min_value=0, step=1, format="%d"),
            "Pr.Unit.":      st.column_config.NumberColumn("Pr.Unit.", min_value=0.0, step=0.01, format="%.2f"),
            "Tabela de IVA": st.column_config.NumberColumn("Tabela de IVA", min_value=0, step=1, format="%d"),
            "Cor":           st.column_config.TextColumn("Cor"),
            "Tamanho":       st.column_config.TextColumn("Tamanho"),
            "Delivery Date": st.column_config.TextColumn("Delivery Date"),
            "Nº SPO":        st.column_config.TextColumn("Nº SPO"),
            "Supplier":      st.column_config.TextColumn("Supplier"),
        },
        key="index_editor"
    )
    st.session_state["index_df"] = edited_df

    valid = edited_df[pd.to_numeric(edited_df["Quant."], errors="coerce").fillna(0) > 0].copy()

    if not valid.empty:
        data_list = []
        for _, row in valid.iterrows():
            qty   = pd.to_numeric(row["Quant."], errors="coerce") or 0
            price = pd.to_numeric(row["Pr.Unit."], errors="coerce") or 0
            data_list.append(make_row(
                ref        = row["Referência"],
                des        = row["Designação"],
                qty        = qty,
                price      = price,
                vat        = row["Tabela de IVA"],
                color      = row["Cor"],
                size       = row["Tamanho"],
                dest       = row["Supplier"],
                spo        = row["Nº SPO"],
                client_date= row["Delivery Date"],
            ))

        df_final = pd.DataFrame(data_list)
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as writer:
            df_final[cols].to_excel(writer, index=False, sheet_name="Index")

        st.download_button(
            f"⬇️ Download PHC Excel ({len(data_list)} rows)",
            out.getvalue(),
            "IMPORT_Index.xlsx",
            key="dl_index"
        )
