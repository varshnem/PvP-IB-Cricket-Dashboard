import pandas as pd

FILE = "PvP IB Cricket Dashboard.xlsx"

df = pd.read_excel(
    FILE,
    sheet_name="Points_Tables",
    header=None
)

pd.set_option("display.max_rows", 100)

print(df.iloc[0:80,0:20])