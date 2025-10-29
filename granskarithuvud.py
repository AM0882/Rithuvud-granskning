import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO

st.title("PDF Text Extractor with Bounding Box")

st.markdown("""
Ladda upp ritningar och exportera info i rithuvud.  
v.1.6
""")

# Constants
MM_TO_PT = 2.83465

# Define bounding boxes in mm from bottom-right corner: (x1, x2, y1, y2)
BOXES_MM = {
    "STATUS": (20, 110, 111, 121),
    "HANDLING": (20, 110, 101, 111),
    "DATUM": (90, 110, 94, 99),
    "ÄNDRING": (20, 90, 94, 99),
    "PROJEKT": (10, 110, 74, 92),
    "KONTAKTPERSON": (60, 110, 47, 52),
    "SKAPAD AV": (10, 60, 47, 52),
    "GODKÄND AV": (60, 110, 40, 45),
    "UPPDRAGSNUMMER": (10, 60, 40, 45),
    "RITNINGSKATEGORI": (28.6, 110, 33, 40),
    "INNEHÅLL": (40, 110, 19, 33),
    "FORMAT": (10, 30, 26, 31),
    "SKALA": (17, 30, 19, 26),
    "NUMMER": (39, 110, 10, 19),
    "BET": (14.7, 30, 10, 19)
}

uploaded_files = st.file_uploader("Upload PDF files", type="pdf", accept_multiple_files=True)

def mm_box_to_pdf_bbox(page_width, page_height, x1_mm, x2_mm, y1_mm, y2_mm):
    x1_pt = page_width - x2_mm * MM_TO_PT
    x2_pt = page_width - x1_mm * MM_TO_PT
    y1_pt = page_height - y2_mm * MM_TO_PT
    y2_pt = page_height - y1_mm * MM_TO_PT
    return (x1_pt, y1_pt, x2_pt, y2_pt)

def extract_boxes(pdf_file, filename):
    extracted_rows = []

    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            page_width = page.width
            page_height = page.height

            row = {"File": filename}

            for field, (x1_mm, x2_mm, y1_mm, y2_mm) in BOXES_MM.items():
                bbox = mm_box_to_pdf_bbox(page_width, page_height, x1_mm, x2_mm, y1_mm, y2_mm)
                cropped = page.within_bbox(bbox)
                text = cropped.extract_text()
                row[field] = text.strip() if text else ""

            extracted_rows.append(row)

    return extracted_rows

if uploaded_files:
    all_data = []

    for file in uploaded_files:
        all_data.extend(extract_boxes(file, file.name))

    df = pd.DataFrame(all_data)

    # Remove 'Text' column if present
    if "Text" in df.columns:
        df.drop(columns=["Text"], inplace=True)

    output = BytesIO()
    df.to_excel(output, index=False, engine='openpyxl')
    output.seek(0)

    st.success("Metadata extracted successfully!")
    st.download_button(
        label="Download Excel file",
        data=output,
        file_name="metadata_boxes.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

