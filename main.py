import pandas as pd
import glob
# glob is a standard module that helps to get all the files in a directory as a list
filepaths=glob.glob("invoices/*.xlsx")

for i in filepaths:
    df=pd.read_excel(i,sheet_name="Sheet 1")
    print(df)