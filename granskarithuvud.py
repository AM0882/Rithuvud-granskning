import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill
import time

st.title("Hämta ut info från rithuvud")

st.markdown("""
Ladda upp ritningar och exportera info i rithuvud. Jämför ritningsnummer med filnamn. Fungerar bara om filer är plottade rätt så rithuvud inte är förskjutet, baserat på ett specifikt projekt iykyk.  
v.1.15
""")

# Constants
MM_TO_PT = 2.83465

# Koordinater för rithuvud, K2/3 har marginaler i nedersta hörnan = 10 och 10mm, K1 har 20 och 20 mm. K12 är anpassad för felplottade 5271
BOXES_K2K3_MM = {
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
    "INNEHÅLL": (57, 110, 19, 33),
    "FORMAT": (10, 30, 26, 31),
    "SKALA": (17, 30, 19, 26),
    "NUMMER": (39, 110, 10, 19),
    "BET": (14.7, 30, 10, 19)
}

# Välj koordinatsystem
coordinate_option = st.selectbox(
    "Välj filstorlek för ritning",
    options=["Helplan", "A1", "A1-5271"]
)

if coordinate_option == "Helplan":
    BOXES_MM = BOXES_K2K3_MM
elif coordinate_option == "A1":
    BOXES_MM = BOXES_K2K3_MM
elif coordinate_option == "A1-5271":
    BOXES_MM = BOXES_K2K3_MM

uploaded_files = st.file_uploader("Ladda upp PDF", type="pdf", accept_multiple_files=True)

status_placeholder = st.empty()

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
# Kör endast när "Starta" trycks
if st.button("Starta") and uploaded_files:
    status_placeholder.info("Körning pågår...")

    all_data = []
    progress_bar = st.progress(0)
    total_files = len(uploaded_files)

    for i, file in enumerate(uploaded_files):
        all_data.extend(extract_boxes(file, file.name))
        progress_bar.progress((i + 1) / total_files)

    df = pd.DataFrame(all_data)

    wb = Workbook()
    ws = wb.active
    ws.title = "Metadata"
    ws.append(df.columns.tolist())

    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

    for index, row in df.iterrows():
        excel_row = [row[col] for col in df.columns]
        ws.append(excel_row)

        file_val = str(row["File"]).strip().lower().replace(".pdf", "")
        nummer_val = str(row["NUMMER"]).strip().lower()

        if file_val != nummer_val:
            cell = ws.cell(row=ws.max_row, column=df.columns.get_loc("NUMMER") + 1)
            cell.fill = red_fill

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    status_placeholder.success("Export och jämförelse färdig!")
    st.download_button(
        label="Ladda ner sammanfattning",
        data=output,
        file_name="metadata_comparison.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


