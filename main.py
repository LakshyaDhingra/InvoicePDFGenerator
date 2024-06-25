import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("Invoices/*.xlsx")

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    # Creation of PDF
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    # Extracting invoice number and date from filename
    filename = Path(filepath).stem
    invoice_number = filename.split("-")[0]
    date = filename.split("-")[1]
    # Setting font of the title invoice number
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=0, h=8, txt=f"Invoice no. {invoice_number}", align="R")
    # Setting font of the date
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=0, h=22, txt=f"Date: {date}", align="R")

    # Generation of PDF
    pdf.output(f"PDFs/{filename}.pdf")


