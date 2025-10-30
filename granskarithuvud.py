
import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from PIL import Image
from streamlit_drawable_canvas import st_canvas

st.title("Interaktiv extraktion från rithuvud")

st.markdown("""
Ladda upp ritningar och välj rutor direkt på första sidan. 
Exportera metadata till Excel med jämförelse av filnamn och ritningsnummer.
V.1.9
""")

uploaded_files = st.file_uploader("Ladda upp PDF", type="pdf", accept_multiple_files=True)

if uploaded_files:
    first_file = uploaded_files[0]

    # Visa första sidan som bild
    with pdfplumber.open(first_file) as pdf:
        first_page = pdf.pages[0]
        page_image = first_page.to_image(resolution=150).original
        page_width = first_page.width
        page_height = first_page.height

    st.image(page_image, caption="Första sidan – rita rutor för extraktion")

    canvas_result = st_canvas(
        fill_color="rgba(255, 165, 0, 0.3)",
        stroke_width=2,
        background_image=page_image,
        update_streamlit=True,
        height=page_image.height,
        width=page_image.width,
        drawing_mode="rect",
        key="canvas"
    )

    if canvas_result.json_data:
        boxes = canvas_result.json_data["objects"]
        st.success(f"{len(boxes)} rutor valda.")

        # Namnge varje ruta
        box_names = []
        for i in range(len(boxes)):
            box_names.append(st.text_input(f"Namn för ruta {i+1}", value=f"Fält_{i+1}"))

        # Konvertera canvas-pixelkoordinater till PDF-punkter
        def canvas_to_pdf_bbox(obj):
            left = obj["left"]
            top = obj["top"]
            width = obj["width"]
            height = obj["height"]
            x1 = left * page_width / page_image.width
            x2 = (left + width) * page_width / page_image.width
            y1 = top * page_height / page_image.height
            y2 = (top + height) * page_height / page_image.height
            return (x1, y1, x2, y2)

        user_boxes = {name: canvas_to_pdf_bbox(obj) for name, obj in zip(box_names, boxes)}

        def extract_boxes(pdf_file, filename):
            extracted_rows = []
            with pdfplumber.open(pdf_file) as pdf:
                for page in pdf.pages:
                    row = {"File": filename}
                    for field, bbox in user_boxes.items():
                        cropped = page.within_bbox(bbox)
                        text = cropped.extract_text()
                        row[field] = text.strip() if text else ""
                    extracted_rows.append(row)
            return extracted_rows

        all_data = []
        progress_bar = st.progress(0)
        total_files = len(uploaded_files)

        for i, file in enumerate(uploaded_files):
            all_data.extend(extract_boxes(file, file.name))
            progress_bar.progress((i + 1) / total_files)

        df = pd.DataFrame(all_data)

        # Skapa Excel-fil
        wb = Workbook()
        ws = wb.active
        ws.title = "Metadata"
        ws.append(df.columns.tolist())

        red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

        for index, row in df.iterrows():
            excel_row = [row[col] for col in df.columns]
            ws.append(excel_row)
            file_val = str(row["File"]).strip().lower().replace(".pdf", "")
            nummer_val = str(row.get("NUMMER", "")).strip().lower()
            if file_val != nummer_val:
                cell = ws.cell(row=ws.max_row, column=df.columns.get_loc("NUMMER") + 1)
                cell.fill = red_fill

        output = BytesIO()
        wb.save(output)
        output.seek(0)

        st.download_button(
            label="Ladda ner sammanfattning",
            data=output,
            file_name="metadata_comparison.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

