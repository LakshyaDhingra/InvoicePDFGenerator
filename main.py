import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("Invoices/*.xlsx")

for filepath in filepaths:
    # Creation of PDF
    pdf2 = FPDF(orientation="P", unit="mm", format="A4")
    pdf2.add_page()
    # Extracting invoice number and date from filename
    filename = Path(filepath).stem
    invoice_number = filename.split("-")[0]
    date = filename.split("-")[1]
    # Setting font of the title invoice
    pdf2.set_font(family="Helvetica", size=34, style="B")
    pdf2.cell(w=0, h=16, txt="Invoice", align="R")
    # Setting font of the invoice number
    pdf2.set_font(family="Helvetica", size=16)
    pdf2.cell(w=0, h=32, txt=f"#{invoice_number}", align="R")
    # Setting font of the date
    pdf2.set_font(family="Times", size=16, style="B")
    pdf2.cell(w=0, h=56, txt=f"Date: {date}", align="R", ln=1)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Creation of headers
    columns = list(df.columns)
    columns = [item.replace("_", " ").title() for item in columns]
    pdf2.set_font(family="Times", size=10, style="B")
    pdf2.set_text_color(r=80, g=80, b=80)
    pdf2.cell(w=30, h=8, txt=(columns[0]), border=1)
    pdf2.cell(w=50, h=8, txt=(columns[1]), border=1)
    pdf2.cell(w=40, h=8, txt=(columns[2]), border=1)
    pdf2.cell(w=30, h=8, txt=(columns[3]), border=1)
    pdf2.cell(w=30, h=8, txt=(columns[4]), border=1, ln=1)

    # Creation of rows
    for index, row in df.iterrows():
        pdf2.set_font(family="Times", size=10)
        pdf2.set_text_color(r=80, g=80, b=80)
        pdf2.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf2.cell(w=50, h=8, txt=str(row["product_name"]), border=1)
        pdf2.cell(w=40, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf2.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf2.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    # Total price calculation and display
    total_sum = df["total_price"].sum()
    pdf2.set_font(family="Times", size=10)
    pdf2.set_text_color(r=80, g=80, b=80)
    pdf2.cell(w=30, h=10, txt=f"The total amount is {total_sum} Rupees")

    # Generation of PDF
    pdf2.output(f"PDFs/{filename}.pdf")

# Text files in PDF
filepaths2 = glob.glob("TxtFiles/*.txt")

pdf2 = FPDF(orientation="P", unit="mm", format="A4")

for filepath2 in filepaths2:
    # Creation of PDF

    pdf2.add_page()

    filename2 = Path(filepath2).stem
    name = filename2.title()

    pdf2.set_font(family="Helvetica", size=16, style="B")
    pdf2.cell(w=0, h=16, txt=name, align="L")

    # Generation of PDF
pdf2.output(f"PDFs/txt_files_output.pdf")

