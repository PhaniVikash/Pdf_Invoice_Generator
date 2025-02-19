import pandas as pd
from fpdf import FPDF
import glob
from pathlib import Path

# glob is a standard module that helps to get all the files in a directory as a list
filepaths=glob.glob("invoices/*.xlsx")

for filepath in filepaths:

    #Create PDF using FPDF file
    pdf = FPDF(orientation="P",unit="mm",format="A4")
    pdf.add_page()

    #Extract filename and date from file
    filename=Path(filepath).stem
    invoice_nr=filename.split("-")[0]
    date=filename.split("-")[1]

    pdf.set_font(family="Times", style="B", size=18)
    pdf.cell(w=0,h=12,txt=f"Invoice Number  : {invoice_nr}",align="L",border=0)

    pdf.set_font(family="Times",style="B",size=18)
    pdf.cell(w=0,h=12,txt=f"Date :{date} ",align="R",border=0,ln=1)
    pdf.ln()

    #Create Column Name with Capitalize
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    col = df.columns
    col=[i.replace("_"," ").title() for i in col]
    pdf.set_font(family="Times", style="B", size=12)
    pdf.set_text_color(0, 0, 0)

    #Add Column Names
    pdf.cell(w=10, h=12, txt="S.No", border=1, align="C")
    pdf.cell(w=25, h=12, txt=f"{col[0]}", border=1, align="C")
    pdf.cell(w=60, h=12, txt=f"{col[1]}", border=1, align="C")
    pdf.cell(w=42, h=12, txt=f"{col[2]}", border=1, align="C")
    pdf.cell(w=30, h=12, txt=f"{col[3]}", border=1, align="C")
    pdf.cell(w=30, h=12, txt=f"{col[4]}", border=1, align="C", ln=1)

    #Add tables
    for index, items in df.iterrows():
        pdf.set_font(family="Times",style="I",size=12)
        pdf.set_text_color(80,80,80)
        pdf.cell(w=10, h=12, txt=str(index+1), border=1,align="C")
        pdf.cell(w=25,h=12,txt=f"{items["product_id"]}",border=1,align="C")
        pdf.cell(w=60, h=12, txt=f"{items["product_name"]}", border=1,align="C")
        pdf.cell(w=42, h=12, txt=f"{items["amount_purchased"]}", border=1,align="C")
        pdf.cell(w=30, h=12, txt=f"{items["price_per_unit"]}", border=1,align="C")
        pdf.cell(w=30, h=12, txt=f"{items["total_price"]}", border=1,align="C", ln=1)

    total=df["total_price"].sum()
    pdf.ln()
    pdf.set_font(family="Times", style="B", size=14)
    pdf.set_text_color(0,0,0)
    pdf.cell(w=0, h=12, txt=f"Total price of the items  : {total}", border=0, align="L", ln=1)

    pdf.set_font(family="Times", style="B", size=20)
    pdf.cell(w=60, h=12, txt=f"Vikash Solutions ", border=0, align="L")
    pdf.image("pythonhow.png",w=10)


    pdf.output(f"PDF's/{filename}.pdf")
