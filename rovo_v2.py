import streamlit as st
import pandas as pd
import pdfplumber
import re
import io

# ── helpers ────────────────────────────────────────────────────────────────────

SIZE_PATTERN = re.compile(r"UK\d+\s*/\s*IT\d+", re.IGNORECASE)
PRICE_PATTERN = re.compile(r"€\s*([\d,\.]+)")
IGNORE_LINES  = {"total qty", "total cost", "sub-total", "first/make", "qty"}

MODEL_TRIGGERS = ["SNW -", "SNM -", "SN -", "LAY "]

def is_model_line(text):
    return any(t in text.upper() for t in MODEL_TRIGGERS)

def extract_price(text):
    """Return first price found in line, as float."""
    m = PRICE_PATTERN.search(text)
    if m:
        return float(m.group(1).replace(",", ""))
    return None

def extract_ship_to(full_text):
    """Pull the destination city/name after 'Ship To:'."""
    m = re.search(r"Ship To:\s*\n(.+)", full_text)
    if m:
        return m.group(1).strip()
    return "Ver PDF"

def clean_color(text, sizes):
    """Remove size tokens, numbers, € and known noise — what's left is the color."""
    noise = {"QTY", "COST", "TOTAL", "FIRST", "MAKE", "DOCKET", "JERSEY",
             "MICRO", "RIB", "SHORT", "SLEEVE", "NECK", "VEST", "HENLEY",
             "COTTON", "BRANDED", "BOXY", "FIT", "T-SHIRT", "-"}
    parts = []
    for tok in text.split():
        up = tok.upper().rstrip(".,")
        if SIZE_PATTERN.match(tok):
            continue
        if up in noise or re.fullmatch(r"[\d,\.€]+", tok) or len(up) <= 1:
            continue
        parts.append(tok)
    return " ".join(parts).strip()


# ── main extractor ─────────────────────────────────────────────────────────────

def extract_studio_nicholson(pdf_file):
    rows = []

    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            full_text = page.extract_text() or ""
            lines = full_text.split("\n")
            destino = extract_ship_to(full_text)

            # ── Parse line by line with a small state machine ──────────────────
            model      = ""       # e.g. "SORIN SNW - 1868 MICRO RIB"
            color      = ""       # e.g. "BLACK"
            sizes      = []       # ["UK4 / IT36", "UK6 / IT38", ...]
            quantities = []       # [11, 9, 8, 6, 3, 0]
            price      = 0.0

            i = 0
            while i < len(lines):
                raw   = lines[i]
                strip = raw.strip()
                up    = strip.upper()

                # Skip total/summary lines
                if any(up.startswith(k) for k in IGNORE_LINES):
                    i += 1
                    continue

                # ── 1. Model line (may span 2 lines) ──────────────────────────
                if is_model_line(strip):
                    # Flush previous product if we have one
                    if model and color and sizes and quantities and price:
                        _flush(rows, model, color, sizes, quantities, price, destino)

                    model      = strip
                    color      = ""
                    sizes      = []
                    quantities = []
                    price      = 0.0

                    # Check if next line is continuation of model (no sizes, no €, not a color-only line)
                    if i + 1 < len(lines):
                        nxt = lines[i + 1].strip()
                        if nxt and not SIZE_PATTERN.search(nxt) and "€" not in nxt and not is_model_line(nxt):
                            # Looks like fabric continuation: "MICRO RIB"
                            if not any(c.isdigit() for c in nxt):
                                model += " " + nxt
                                i += 1

                    i += 1
                    continue

                # ── 2. Size header line ────────────────────────────────────────
                if SIZE_PATTERN.search(strip):
                    sizes = SIZE_PATTERN.findall(strip)
                    # Normalize: "UK4 / IT36" → "UK4/IT36" or keep original
                    sizes = [s.replace(" ", "") for s in sizes]
                    i += 1
                    continue

                # ── 3. Quantity line (all tokens are digits or "-") ────────────
                tokens = strip.split()
                if tokens and all(re.fullmatch(r"\d+|-", t) for t in tokens) and len(tokens) >= 2:
                    quantities = [int(t) if t.isdigit() else 0 for t in tokens]
                    i += 1
                    continue

                # ── 4. Price line ──────────────────────────────────────────────
                p = extract_price(strip)
                if p is not None and ("cost" in up or "€" in strip) and "total" not in up:
                    price = p
                    # Try to flush if we have everything
                    if model and color and sizes and quantities:
                        _flush(rows, model, color, sizes, quantities, price, destino)
                        color      = ""
                        sizes      = []
                        quantities = []
                        price      = 0.0
                    i += 1
                    continue

                # ── 5. Color line (after model, before sizes) ──────────────────
                # Heuristic: non-empty, no sizes, no €, no digits-only tokens, not noise
                if model and not sizes and strip and "€" not in strip:
                    candidate = clean_color(strip, sizes)
                    if candidate and not any(c.isdigit() for c in candidate):
                        color = candidate

                i += 1

    return rows


def _flush(rows, model, color, sizes, quantities, price, destino):
    """Append one row per size that has quantity > 0."""
    for idx, size in enumerate(sizes):
        qty = quantities[idx] if idx < len(quantities) else 0
        if qty > 0:
            rows.append({
                "Referência":      "",
                "Designação":      model,
                "Quant.":          qty,
                "Pr.Unit.":        price,
                "Pr.Unit.Moeda":   0,
                "Tabela de IVA":   4,
                "Cor":             color,
                "Tamanho":         size,
                "TOTAL":           round(qty * price, 2),
                "Destino":         destino,
                "CPO":             "",
            })


# ── Streamlit UI ───────────────────────────────────────────────────────────────

st.set_page_config(page_title="ROVO - Conversor Universal", page_icon="🚀", layout="wide")
st.sidebar.title("🚀 MENU ROVO")
cliente  = st.sidebar.selectbox("Selecione o Cliente", ["Studio Nicholson", "Stussy", "Supreme"])
st.title(f"📦 Conversor: {cliente}")

arquivo = st.file_uploader("Submeter ficheiro", type=["xlsx", "pdf"])

if arquivo:
    try:
        lista_dados = []

        if arquivo.name.endswith(".xlsx"):
            pass  # Lógica Excel existente mantida aqui

        elif arquivo.name.endswith(".pdf") and cliente == "Studio Nicholson":
            with pdfplumber.open(arquivo) as pdf:
            for i, page in enumerate(pdf.pages):
            st.subheader(f"Página {i+1} — Texto")
            st.text(page.extract_text())
            st.subheader(f"Página {i+1} — Palavras")
            st.write(page.extract_words())

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
        st.exception(e)  # Mostra o traceback completo durante desenvolvimento
