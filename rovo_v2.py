elif arquivo.name.endswith(".pdf") and cliente == "Studio Nicholson":
    with pdfplumber.open(arquivo) as pdf:
        for i, page in enumerate(pdf.pages):
            st.subheader(f"Página {i+1} — Texto raw")
            st.text(page.extract_text())
