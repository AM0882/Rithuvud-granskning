
import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
from io import BytesIO

st.title("PDF Text Extractor")

st.markdown("""
Ladda upp ritningar och exportera info i rithuvud.  
v.1.1
""")

uploaded_files = st.file_uploader("Upload PDF files", type="pdf", accept_multiple_files=True)

def extract_bottom_right_text(pdf_file, filename):
    doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
    extracted_text = []

    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        blocks = page.get_text("blocks")

        # Sort blocks by their position (bottom right)
        blocks_sorted = sorted(blocks, key=lambda b: (-b[1], -b[0]))

        if blocks_sorted:
            bottom_right_block = blocks_sorted[0]
            extracted_text.append({
                "File": filename,
                "Page": page_num + 1,
                "Text": bottom_right_block[4]
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
        file_name="bottom_right_text.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
