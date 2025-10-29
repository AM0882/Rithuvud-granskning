import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO

st.title("PDF Text Extractor with Bounding Box")

st.markdown("""
Ladda upp ritningar och exportera info i rithuvud.  
v.1.5 
""")

# Constants
MM_TO_PT = 2.83465
BOX_WIDTH_PT = 114 * MM_TO_PT
BOX_HEIGHT_PT = 120 * MM_TO_PT

# Metadata fields to extract
METADATA_FIELDS = [
    "STATUS", "HANDLING", "DATUM", "ÄNDRING", "PROJEKT", "ANSVARIG PART",
    "KONTAKTPERSON", "SKAPAD AV", "GODKÄND AV", "UPPDRAGSNUMMER",
    "RITNINGSKATEGORI", "INNEHÅLL", "FORMAT", "SKALA", "NUMMER", "BET"
]

st.title("PDF Metadata Extractor with pdfplumber")

uploaded_files = st.file_uploader("Upload PDF files", type="pdf", accept_multiple_files=True)

def extract_metadata(text):
    metadata = {field: "" for field in METADATA_FIELDS}
    lines = text.splitlines()
    for line in lines:
        for field in METADATA_FIELDS:
            if line.strip().upper().startswith(field.upper()):
                parts = line.split(":", 1)
                if len(parts) == 2:
                    metadata[field] = parts[1].strip()
                else:
                    metadata[field] = line.replace(field, "").strip()
    return metadata
def extract_text_from_area(pdf_file, filename):
    extracted_rows = []

    with pdfplumber.open(pdf_file) as pdf:
        for page_num, page in enumerate(pdf.pages):
            page_width = page.width
            page_height = page.height

            # Define bottom-right box
            x0 = page_width - BOX_WIDTH_PT
            y0 = page_height - BOX_HEIGHT_PT
            x1 = page_width
            y1 = page_height

            cropped = page.within_bbox((x0, y0, x1, y1))
            text = cropped.extract_text()

            if text:
                metadata = extract_metadata(text)
                metadata["File"] = filename
                extracted_rows.append(metadata)

    return extracted_rows

if uploaded_files:
    all_data = []

    for file in uploaded_files:
        all_data.extend(extract_text_from_area(file, file.name))

    df = pd.DataFrame(all_data)

    output = BytesIO()
    df.to_excel(output, index=False, engine='openpyxl')
    output.seek(0)

    st.success("Metadata extracted successfully!")
    st.download_button(
        label="Download Excel file",
        data=output,
        file_name="metadata_only.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
