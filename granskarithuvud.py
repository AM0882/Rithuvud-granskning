import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO

st.title("PDF Text Extractor with Bounding Box")

st.markdown("""
Ladda upp ritningar och exportera info i rithuvud.  
v.1.4 – med pdfplumber och bounding box
""")

# Convert mm to points
BOX_WIDTH_PT = 114 * 2.83465
BOX_HEIGHT_PT = 120 * 2.83465

st.title("PDF Text Extractor with pdfplumber")

uploaded_files = st.file_uploader("Upload PDF files", type="pdf", accept_multiple_files=True)

def extract_text_from_area(pdf_file, filename):
    extracted_text = []

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
                extracted_text.append({
                    "File": filename,
                    "Text": text.strip()
                })

    return extracted_text

if uploaded_files:
    all_text_data = []

    for file in uploaded_files:
        all_text_data.extend(extract_text_from_area(file, file.name))

    df = pd.DataFrame(all_text_data)

    output = BytesIO()
    df.to_excel(output, index=False, engine='openpyxl')
    output.seek(0)

    st.success("Text extracted successfully!")
    st.download_button(
        label="Download Excel file",
        data=output,
        file_name="bottom_right_text.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
