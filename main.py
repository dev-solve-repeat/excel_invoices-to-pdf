
import pandas as pd
import glob

filepath = glob.glob("invoices/*.xlsx")
print(filepath)

for path in filepath:
    df = pd.read_excel(path, sheet_name="Sheet 1")
    print(df)