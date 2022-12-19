
import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepath = glob.glob("invoices/*.xlsx")
print(filepath)

for path in filepath:

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    filename = Path(path).stem
    invoice_no, date = filename.split("-")

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice Nr - {invoice_no}", ln=1)

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Date - {date}", ln=1)

    df = pd.read_excel(path, sheet_name="Sheet 1")

    #Add a header
    col = df.columns
    col = [items.replace("_", " ").title() for items in col]
    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=10, txt=col[0], border=1)
    pdf.cell(w=50, h=10, txt=col[1], border=1)
    pdf.cell(w=40, h=10, txt=col[2], border=1)
    pdf.cell(w=30, h=10, txt=col[3], border=1)
    pdf.cell(w=30, h=10, txt=col[4], border=1, ln=1)

    #Add row to the table
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=10, txt=str(row["product_id"]), border=1) #the input to txt should be string, no int
        pdf.cell(w=50, h=10, txt=str(row["product_name"]), border=1)
        pdf.cell(w=40, h=10, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=10, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=10, txt=str(row["total_price"]), border=1, ln=1)

    total_sum = df["total_price"].sum()
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=10, txt=" ", border=1)  # the input to txt should be string, no int
    pdf.cell(w=50, h=10, txt=" ", border=1)
    pdf.cell(w=40, h=10, txt=" ", border=1)
    pdf.cell(w=30, h=10, txt=" ", border=1)
    pdf.cell(w=30, h=10, txt=str(total_sum), border=1, ln=1)

    #To add last line
    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=0, h=10, txt=f"The total price is {total_sum}", ln=1)

    #To add the image
    pdf.set_font(family="Times", size=16, style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=10, txt="PythonHow")
    pdf.image("pythonhow.png", w=10)

    pdf.output(f"PDF/{filename}.pdf")