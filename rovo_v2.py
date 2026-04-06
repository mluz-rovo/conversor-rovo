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


# ---------------------------------------------------------------------------
# STUDIO NICHOLSON — PASS 1: Raw extraction
# Returns a list of page-level dicts with all the raw data we need.
# ---------------------------------------------------------------------------
SIZE_REFS = ["XXS", "XS", "S", "M", "L", "XL", "XXL",
             "UK4", "UK6", "UK8", "UK10", "UK12", "UK14"]

SKIP_LINES = ["TOTAL QTY", "FIRST/MAKE", "SUB-TOTAL", "QTY", "COST", "TOTAL"]
MODEL_PREFIXES = ["SNW -", "SNM -", "SN -", "LAY "]
JUNK_WORDS = {
    "JERSEY", "MICRO", "RIB", "SHORT", "SLEEVE", "NECK", "VEST", "HENLEY",
    "COTTON", "BRANDED", "BOXY", "FIT", "T-SHIRT", "QTY", "COST", "TOTAL",
    "FIRST", "MAKE"
}


def sn_extract_raw(pdf_file) -> list[dict]:
    """
    Pass 1: Walk every page and return a list of raw 'events':
      - type 'destination'  → {'dest': str}
      - type 'model'        → {'model': str}
      - type 'price_row'    → {'color': str, 'unit_price': float,
                               'size_quantities': {size_label: qty}}
    No business logic here — just faithful extraction.
    """
    events = []

    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            words   = page.extract_words(keep_blank_chars=False)
            text    = page.extract_text() or ""
            lines   = text.split("\n")

            # --- Destination (Ship To) ---
            if "Ship To:" in text:
                after      = text.split("Ship To:", 1)[1]
                dest_lines = [l.strip() for l in after.split("\n") if l.strip()]
                # First non-empty line after "Ship To:" is usually the company name
                if dest_lines:
                    events.append({"type": "destination", "dest": dest_lines[0]})

            # --- Size column map: size label → x-centre on this page ---
            size_map = _build_size_map(words)

            # --- Walk lines ---
            for i, line in enumerate(lines):
                l_up = line.upper().strip()

                # Lines we never want
                if not l_up or any(skip in l_up for skip in SKIP_LINES):
                    continue

                # Model header line
                if any(l_up.startswith(pfx.upper()) for pfx in MODEL_PREFIXES):
                    model = re.split(r"Qty|Cost|Total|First", line, flags=re.I)[0].strip()
                    events.append({"type": "model", "model": model})
                    continue

                # Price / colour row (contains a € symbol)
                if "€" in line:
                    event = _parse_price_row(line, lines, i, words, page, size_map)
                    if event:
                        events.append(event)

    return events


def _build_size_map(words: list[dict]) -> list[dict]:
    """Return list of {size, center_x} for every recognised size label on the page."""
    size_map = []
    for w in words:
        t = w["text"].upper().strip()
        if t in SIZE_REFS or any(ref in t and "/" in t for ref in SIZE_REFS):
            cx = (w["x0"] + w["x1"]) / 2
            size_map.append({"size": t, "center_x": cx, "x1": w["x1"]})
    return size_map


def _parse_price_row(line, lines, line_idx, words, page, size_map) -> dict | None:
    """
    Extract colour, unit price and {size: qty} from a single price row.
    Returns a 'price_row' event dict or None if nothing useful found.
    """
    # Unit price — first €-prefixed decimal on the line
    prices = re.findall(r"[\d]+[.,]\d{2}", line)
    if not prices:
        return None
    unit_price = float(prices[0].replace(",", "."))

    # Colour — tokens that aren't numbers, junk words, or the € symbol
    color_tokens = [
        pt for pt in line.split()
        if pt.upper().replace(",", "").replace(".", "") not in JUNK_WORDS
        and not re.match(r"^[\d.,€]+$", pt)
        and len(pt) > 1
    ]
    color = " ".join(color_tokens[:2])

    # Find the y-coordinate of this line by matching against word positions
    line_text = line.strip()
    y_ref = None
    for w in words:
        if w["text"] in line_text:
            y_ref = w["top"]
            break
    if y_ref is None:
        return None

    # Match quantities to sizes by proximity on the x-axis
    size_quantities = {}
    x_max = max((m["x1"] for m in size_map), default=0)

    for size_info in size_map:
        for w in words:
            if (abs(w["top"] - y_ref) < 10
                    and abs((w["x0"] + w["x1"]) / 2 - size_info["center_x"]) < 15
                    and w["text"].isdigit()
                    and w["x1"] <= x_max + 10):
                qty = int(w["text"])
                if qty > 0:
                    size_quantities[size_info["size"]] = qty
                break  # one word per size column

    if not size_quantities:
        return None

    return {
        "type": "price_row",
        "color": color,
        "unit_price": unit_price,
        "size_quantities": size_quantities,
    }


# ---------------------------------------------------------------------------
# STUDIO NICHOLSON — PASS 2: Transform raw events → flat row dicts
# ---------------------------------------------------------------------------
def sn_transform(events: list[dict]) -> list[dict]:
    """
    Pass 2: Walk the event stream and build the final flat rows.
    State machine: track current destination, current model.
    """
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
                    "Reference":            "",
                    "Designation":          current_model,
                    "Qty":                  qty,
                    "Unit Price":           ev["unit_price"],
                    "Unit Price Currency":  0,
                    "VAT Table":            4,
                    "Color":                ev["color"],
                    "Size":                 size,
                    "TOTAL":                qty * ev["unit_price"],
                    "Destination":          current_dest,
                    "CPO No.":              "",
                    "SPO No.":              "",
                    "Supplier Unit Value":  "",
                    "Total Supplier":       "",
                })

    return rows


# ---------------------------------------------------------------------------
# DEBUG HELPER — renders the raw event stream in the UI
# ---------------------------------------------------------------------------
def show_debug(events: list[dict]):
    with st.expander("🔍 Debug: Raw extracted events (Pass 1)", expanded=False):
        for i, ev in enumerate(events):
            st.write(f"**{i}** — `{ev['type']}`", ev)


# ---------------------------------------------------------------------------
# MAIN APP
# ---------------------------------------------------------------------------
if uploaded_file:
    try:
        data_list = []

        # ── STUSSY ──────────────────────────────────────────────────────────
        if uploaded_file.name.endswith(".xlsx") and client == "Stussy":
            xl = pd.ExcelFile(uploaded_file, engine="openpyxl")
            sheet_name = "Sheet1" if "Sheet1" in xl.sheet_names else xl.sheet_names[0]
            df = xl.parse(sheet_name, header=None)

            for i, row in df.iloc[1:].iterrows():
                if len(row) >= 14:
                    q = pd.to_numeric(row[12], errors="coerce")
                    p_raw = row[13]
                    if isinstance(p_raw, str):
                        p = pd.to_numeric(re.sub(r"[^\d\.]", "", p_raw.replace(",", ".")), errors="coerce")
                    else:
                        p = pd.to_numeric(p_raw, errors="coerce")

                    if q and q > 0:
                        data_list.append({
                            "Reference":            "",
                            "Designation":          row[8] if len(row) > 8 else "",
                            "Qty":                  q,
                            "Unit Price":           0,
                            "Unit Price Currency":  p if pd.notna(p) else 0,
                            "VAT Table":            4,
                            "Color":                row[7] if len(row) > 7 else "",
                            "Size":                 row[9] if len(row) > 9 else "",
                            "TOTAL":                q * (p if pd.notna(p) else 0),
                            "Destination":          row[4] if len(row) > 4 else "General",
                            "CPO No.":              "",
                            "SPO No.":              "",
                            "Supplier Unit Value":  "",
                            "Total Supplier":       "",
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
                                data_list.append({
                                    "Reference":            ref_manual,
                                    "Designation":          des_manual,
                                    "Qty":                  q,
                                    "Unit Price":           0,
                                    "Unit Price Currency":  p if pd.notna(p) else 0,
                                    "VAT Table":            4,
                                    "Color":                df.iloc[i, 6],
                                    "Size":                 s_name,
                                    "TOTAL":                q * (p if pd.notna(p) else 0),
                                    "Destination":          dest,
                                    "CPO No.":              "",
                                    "SPO No.":              "",
                                    "Supplier Unit Value":  "",
                                    "Total Supplier":       "",
                                })

        # ── STUDIO NICHOLSON ─────────────────────────────────────────────────
        elif uploaded_file.name.endswith(".pdf") and client == "Studio Nicholson":
            events    = sn_extract_raw(uploaded_file)
            show_debug(events)          # collapsible debug panel — remove when stable
            data_list = sn_transform(events)

        # ── OUTPUT ───────────────────────────────────────────────────────────
        if data_list:
            df_final = pd.DataFrame(data_list).drop_duplicates()
            cols = [
                "Reference", "Designation", "Qty", "Unit Price",
                "Unit Price Currency", "VAT Table", "Color", "Size",
                "TOTAL", "Destination", "CPO No.", "SPO No.",
                "Supplier Unit Value", "Total Supplier",
            ]

            out = io.BytesIO()
            with pd.ExcelWriter(out, engine="openpyxl") as writer:
                for dest in df_final["Destination"].unique():
                    safe_name = re.sub(r"[\[\]*:?/\\]", "", str(dest))[:31]
                    df_dest = df_final[df_final["Destination"] == dest]
                    df_dest[cols].to_excel(writer, index=False, sheet_name=safe_name)

            st.success(f"✅ Conversion complete! {len(data_list)} rows generated.")
            st.dataframe(df_final[cols].head(50))   # quick preview
            st.download_button("⬇️ Download PHC Excel", out.getvalue(), f"IMPORT_{client}.xlsx")
        else:
            st.warning("No data found in the file.")

    except Exception as e:
        st.error(f"Error during processing: {e}")
        st.exception(e)   # full traceback visible in dev; remove in production
