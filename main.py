import pandas as pd
from fpdf import FPDF
import glob
from pathlib import Path

# glob is a standard module that helps to get all the files in a directory as a list
filepaths=glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    df=pd.read_excel(filepath,sheet_name="Sheet 1")
    pdf = FPDF(orientation="P",unit="mm",format="A4")
    pdf.add_page()
    filename=Path(filepath).stem
    invoice_nr=filename.split("-")[0]
    pdf.set_font(family="Times", style="B", size=18)
    pdf.cell(w=0,h=12,txt=f"Invoice Number  : {invoice_nr}",align="L",border=0)
    pdf.output(f"PDF's/{filename}.pdf")
