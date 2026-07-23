import pandas as pd

file = "PvP IB Cricket Dashboard.xlsx"

xls = pd.ExcelFile(file)

print(xls.sheet_names)