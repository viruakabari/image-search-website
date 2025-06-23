
import os
import pandas as pd
import barcode
from barcode.writer import ImageWriter
from fpdf import FPDF

input_file = "Type 1.xlsx"
output_file = "output_data.pdf"

barcode_dir = "barcodes"
os.makedirs(barcode_dir, exist_ok=True)

df = pd.read_excel(input_file)

pdf = FPDF()
pdf.set_auto_page_break(auto=True, margin=15)

pdf.add_page()
pdf.set_font("Arial", size=8)

columns = df.columns.tolist()
columns.remove('barcode_number') 

for index, row in df.iterrows():
    for col in columns:
        if col == 'retail_price':
            barcode_data = str(row['barcode_number']).strip()

            if len(barcode_data) < 12:
                barcode_data = barcode_data.zfill(12) 

            try:
                barcode_img = barcode.get('ean13', barcode_data, writer=ImageWriter())

                barcode_path = os.path.join(barcode_dir, f"barcode_{barcode_data}")

                barcode_img.save(barcode_path)

                barcode_path_with_extension = f"{barcode_path}.png"

                if os.path.exists(barcode_path_with_extension):
                    print(f"Barcode image saved at: {barcode_path_with_extension}")

                    pdf.image(barcode_path_with_extension, x=3, y=pdf.get_y(), w=60, h=5) 
                    pdf.set_y(pdf.get_y() + 5) 
                else:
                    print(f"Error: Barcode image not saved at {barcode_path_with_extension}")
            except Exception as e:
                print(f"Error generating barcode for {barcode_data}: {e}")
                continue

        else:
            header_data = f"{row[col]}"

            pdf.cell(140, 4, header_data, border=0)
            pdf.ln()

    pdf.ln(10)

try:
    pdf.output(output_file)
    print(f"PDF created successfully with barcodes: {output_file}")
except Exception as e:
    print(f"Error creating PDF: {e}")