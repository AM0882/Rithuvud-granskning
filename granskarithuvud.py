import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO

st.title("PDF Text Extractor with Bounding Box")

st.markdown("""
Ladda upp ritningar och exportera info i rithuvud.  
v.1.3 – med pdfplumber och bounding box
""")

uploaded_files = st.file_uploader("Upload PDF files", type="pdf", accept_multiple_files=True)

def extract_bottom_right_text(pdf_file, filename):
    extracted_text = []

    with pdfplumber.open(pdf_file) as pdf:
        for page_num, page in enumerate(pdf.pages):
            width = page.width
            height = page.height
            # Define bounding box: bottom 25% and right 20%
            left = width * 0.80
            top = height * 0.75
            right = width
            bottom = height

            bbox = (left, top, right, bottom)
            cropped = page.within_bbox(bbox)
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
        all_text_data.extend(extract_bottom_right_text(file, file.name))

    df = pd.DataFrame(all_text_data)

    output = BytesIO()
    df.to_excel(output, index=False, engine='openpyxl')
    output.seek(0)

    st.success("Text extracted successfully!")
    st.download_button(
        label="Download Excel file",
        data=output,
        file_name="pdfplumber_bottom_right_text.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
