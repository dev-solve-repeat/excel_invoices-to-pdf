
import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepath = glob.glob("invoices/*.xlsx")
print(filepath)

for path in filepath:
    df = pd.read_excel(path, sheet_name="Sheet 1")

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    filename = Path(path).stem
    invoice_no, date = filename.split("-")

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice Nr - {invoice_no}", ln=1)

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Date - {date}")

    pdf.output(f"PDF/{filename}.pdf")