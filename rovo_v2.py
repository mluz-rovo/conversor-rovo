import streamlit as st
import pandas as pd
import io
import pdfplumber
import re

st.set_page_config(page_title="ROVO - Universal Converter", page_icon="🚀", layout="wide")

st.sidebar.title("🚀 ROVO MENU")
client = st.sidebar.selectbox("Select Client", ["Stussy", "Supreme", "Studio Nicholson"])

# --- MANUAL INPUTS FOR SUPREME ONLY ---
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

if uploaded_file:
    try:
        data_list = []

        if uploaded_file.name.endswith('.xlsx'):
            xl = pd.ExcelFile(uploaded_file, engine='openpyxl')
            
            # --- STUSSY LOGIC ---
            if client == "Stussy":
                # Prioritize 'Sheet1' or first available sheet
                sheet_name = "Sheet1" if "Sheet1" in xl.sheet_names else xl.sheet_names[0]
                df = xl.parse(sheet_name, header=None)
                
                for i, row in df.iloc[1:].iterrows():
                    # Check if row has enough columns (up to R / index 17)
                    if len(row) >= 18:
                        q_raw = row[12] # Column M (Qty)
                        p_raw = row[17] # Column R (Unit Price Currency)
                        
                        q = pd.to_numeric(q_raw, errors='coerce')
                        
                        # Clean Price format
                        if isinstance(p_raw, str):
                            p_clean = re.sub(r'[^\d\.]', '', p_raw.replace(',', '.'))
                            p = pd.to_numeric(p_clean, errors='coerce')
                        else:
                            p = pd.to_numeric(p_raw, errors='coerce')
                        
                        if q and q > 0:
                            p_val = p if pd.notna(p) else 0
                            data_list.append({
                                'Reference': "", 
                                'Designation': row[8] if len(row) > 8 else "", # Column I
                                'Qty': q, 
                                'Unit Price': 0, 
                                'Unit Price Currency': p_val, 
                                'VAT Table': 4, 
                                'Color': row[7] if len(row) > 7 else "", # Column H
                                'Size': row[9] if len(row) > 9 else "", # Column J
                                'TOTAL': q * p_val, 
                                'Destination': row[4] if len(row) > 4 else "General", # Column E
                                'CPO No.': "",
                                'SPO No.': "",
                                'Supplier Unit Value': "",
                                'Total Supplier': ""
                            })

            # --- SUPREME LOGIC ---
            elif client == "Supreme":
                for sheet in xl.sheet_names:
                    if "TOTAL" in sheet.upper(): continue
                    df = xl.parse(sheet, header=None)
                    sizes = {c: str(df.iloc[14, c]) for c in range(9, 16) if pd.notna(df.iloc[14, c])}
                    
                    for start in range(16, len(df), 14):
                        dest = str(df.iloc[start, 0]).strip()
                        if not dest or dest == "nan": dest = "General"
                        
                        for i in range(start + 1, start + 13):
                            if i >= len(df) or pd.isna(df.iloc[i, 6]): continue
                            p = pd.to_numeric(df.iloc[i, 17], errors='coerce')
                            for c_idx, s_name in sizes.items():
                                q = pd.to_numeric(df.iloc[i, c_idx], errors='coerce')
                                if q and q > 0:
                                    p_val = p if pd.notna(p) else 0
                                    data_list.append({
                                        'Reference': ref_manual, 
                                        'Designation': des_manual, 
                                        'Qty': q, 
                                        'Unit Price': 0, 
                                        'Unit Price Currency': p_val, 
                                        'VAT Table': 4, 
                                        'Color': df.iloc[i, 6], 
                                        'Size': s_name, 
                                        'TOTAL': q * p_val, 
                                        'Destination': dest, 
                                        'CPO No.': "",
                                        'SPO No.': "",
                                        'Supplier Unit Value': "",
                                        'Total Supplier': ""
                                    })

        # --- STUDIO NICHOLSON LOGIC ---
        elif uploaded_file.name.endswith('.pdf') and client == "Studio Nicholson":
            with pdfplumber.open(uploaded_file) as pdf:
                size_refs = ["XXS", "XS", "S", "M", "L", "XL", "XXL", "UK4", "UK6", "UK8", "UK10", "UK12", "UK14"]
                junk_words = ["JERSEY", "MICRO", "RIB", "SHORT", "SLEEVE", "NECK", "VEST", "HENLEY", "COTTON", "BRANDED", "BOXY", "FIT", "T-SHIRT", "QTY", "COST", "TOTAL", "FIRST", "MAKE"]
                current_model = ""
                final_dest = "See PDF"
                for page in pdf.pages:
                    words = page.extract_words()
                    text = page.extract_text()
                    if not text: continue
                    lines = text.split('\n')
                    ship_match = re.search(r"Ship To:\s*(.*)", text, re.IGNORECASE)
                    if ship_match:
                        ship_lines = text.split("Ship To:")[1].split('\n')
                        final_dest = ship_lines[1].strip() if len(ship_lines) > 1 else final_dest
                    size_map = []
                    x_max = 0
                    for w in words:
                        t_up = w['text'].upper().strip()
                        if any(t == t_up or (t in t_up and "/" in t_up) for t in size_refs):
                            size_map.append({'size': t_up, 'center': (w['x0']+w['x1'])/2, 'x1': w['x1']})
                            if w['x1'] > x_max: x_max = w['x1']
                    for i, line in enumerate(lines):
                        l_up = line.upper()
                        if any(x in l_up for x in ["TOTAL QTY", "FIRST/MAKE", "SUB-TOTAL"]): continue
                        if any(x in l_up for x in ["SNW -", "SNM -", "SN -", "LAY "]):
                            current_model = re.split(r"Qty|Cost|Total|First", line, flags=re.I)[0].strip()
                            continue
                        if "€" in line:
                            prices = re.findall(r"(\d+[\.,]\d{2})", line)
                            u_price = float(prices[0].replace(',', '')) if prices else 0
                            color_parts = [pt for pt in line.split() if pt.upper().replace(',','').replace('.','') not in junk_words and not pt.isdigit() and "€" not in pt and len(pt) > 2]
                            final_color = " ".join(color_parts[:2])
                            y_ref = next((pw['top'] for pw in words if pw['text'] == "€" and abs(pw['top'] - page.extract_text_lines()[i]['top']) < 20), None)
                            if y_ref is None: continue
                            for m in size_map:
                                for pdw in words:
                                    if abs(pdw['top'] - y_ref) < 10 and abs(((pdw['x0'] + pdw['x1']) / 2) - m['center']) < 12:
                                        if pdw['text'].isdigit() and pdw['x1'] <= (x_max + 10):
                                            q_num = int(pdw['text'])
                                            if q_num > 0:
                                                data_list.append({
                                                    'Reference': "", 'Designation': current_model, 'Qty': q_num, 
                                                    'Unit Price': u_price, 'Unit Price Currency': 0, 'VAT Table': 4, 
                                                    'Color': final_color, 'Size': m['size'], 'TOTAL': q_num * u_price, 
                                                    'Destination': final_dest, 'CPO No.': "", 
                                                    'SPO No.': "", 'Supplier Unit Value': "", 'Total Supplier': ""
                                                })

        # --- FINAL FILE GENERATION ---
        if data_list:
            df_final = pd.DataFrame(data_list).drop_duplicates()
            cols = ['Reference', 'Designation', 'Qty', 'Unit Price', 'Unit Price Currency', 'VAT Table', 'Color', 'Size', 'TOTAL', 'Destination', 'CPO No.', 'SPO No.', 'Supplier Unit Value', 'Total Supplier']
            
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine='openpyxl') as writer:
                for dest in df_final['Destination'].unique():
                    sheet_name = str(dest).replace('[','').replace(']','').replace('*','').replace(':','').replace('?','').replace('/','').replace('\\','')[:31]
                    df_dest = df_final[df_final['Destination'] == dest]
                    df_dest[cols].to_excel(writer, index=False, sheet_name=sheet_name)
            
            st.success(f"✅ Success! Data fetched from Column R for Stussy prices.")
            st.download_button("⬇️ Download PHC Excel", out.getvalue(), f"IMPORT_{client}.xlsx")
        else:
            st.warning("No valid data found. Check your file columns.")

    except Exception as e:
        st.error(f"Error: {e}")
