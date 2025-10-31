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
v.1.19
""")

# Constants
MM_TO_PT = 2.83465

# Koordinater för rithuvud
BOXES_K2K3_MM = {
    "STATUS": (20, 110, 109, 122),
    "HANDLING": (20, 110, 100, 112),
    "DATUM": (89, 110, 94, 100),
    "ÄNDRING": (10, 90, 94, 100),
    "PROJEKT": (10, 112, 73, 93),
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

BOXES_K1_MM = {
    "STATUS": (30, 120, 121, 131),
    "HANDLING": (30, 120, 111, 121),
    "DATUM": (100, 120, 104, 109),
    "ÄNDRING": (30, 100, 104, 109),
    "PROJEKT": (20, 120, 84, 102),
    "KONTAKTPERSON": (70, 120, 57, 62),
    "SKAPAD AV": (20, 70, 57, 62),
    "GODKÄND AV": (70, 120, 50, 55),
    "UPPDRAGSNUMMER": (20, 70, 50, 55),
    "RITNINGSKATEGORI": (38.6, 120, 43, 50),
    "INNEHÅLL": (67, 120, 29, 43),
    "FORMAT": (28.5, 41, 35, 42),
    "SKALA": (28.5, 41, 28, 36),
    "NUMMER": (49, 120, 18, 29.5),
    "BET": (28.5, 41, 18, 29.6)
}

BOXES_K12_MM = {
    "STATUS": (29.8, 119.8, 123, 133),
    "HANDLING": (29.8, 119.8, 113, 123),
    "DATUM": (93.5, 119.8, 106, 111),
    "ÄNDRING": (29.8, 93.5, 106, 111),
    "PROJEKT": (29, 119.8, 86, 106),
    "KONTAKTPERSON": (69.8, 119.8, 59, 64),
    "SKAPAD AV": (19.8, 69.8, 59, 64),
    "GODKÄND AV": (69.8, 119.8, 52, 57),
    "UPPDRAGSNUMMER": (19.8, 69.8, 52, 57),
    "RITNINGSKATEGORI": (38.4, 119.8, 45, 52),
    "INNEHÅLL": (66.8, 119.8, 31, 45),
    "FORMAT": (28.3, 40.8, 37, 44),
    "SKALA": (30.9, 40.8, 30, 38),
    "NUMMER": (48.8, 119.8, 20, 31.5),
    "BET": (28.3, 40.8, 20, 31.6)
}

# Välj koordinatsystem
coordinate_option = st.selectbox(
    "Välj filstorlek för ritning",
    options=["Helplan", "A1", "A1-5271"]
)

if coordinate_option == "Helplan":
    BOXES_MM = BOXES_K2K3_MM
elif coordinate_option == "A1":
    BOXES_MM = BOXES_K1_MM
elif coordinate_option == "A1-5271":
    BOXES_MM = BOXES_K12_MM

uploaded_files = st.file_uploader("Ladda upp PDF", type="pdf", accept_multiple_files=True)

status_placeholder = st.empty()

def mm_box_to_pdf_bbox(page_width, page_height, x1_mm, x2_mm, y1_mm, y2_mm):
    # Räkna från nedre högra hörnet
    x1_pt = page_width - x1_mm * MM_TO_PT
    x2_pt = page_width - x2_mm * MM_TO_PT
    y1_pt = page_height - y2_mm * MM_TO_PT
    y2_pt = page_height - y1_mm * MM_TO_PT

    # Säkerställ att x0 < x1 och y0 < y1
    x0 = min(x1_pt, x2_pt)
    x1 = max(x1_pt, x2_pt)
    y0 = min(y1_pt, y2_pt)
    y1 = max(y1_pt, y2_pt)

    return (x0, y0, x1, y1)

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

