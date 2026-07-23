import pandas as pd

file = "PvP IB Cricket Dashboard.xlsx"

df = pd.read_excel(
    file,
    sheet_name="Points_Tables",
    header=3
)

print(df.head(15))