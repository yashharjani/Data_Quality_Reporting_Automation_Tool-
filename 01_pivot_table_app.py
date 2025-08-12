# Automate with Python – Full Course for Beginners
# https://youtu.be/PXMJ6FS7llk?si=bJs8LTbxuvonaIch

import pandas as pd

df = pd.read_excel("db/Tablero DB_240210.xlsx")

df = df[["Sigla Esp", "Tipo HH", "Horas"]]
# print(df)

pivot_table = df.pivot_table(
    index="Tipo HH", columns="Sigla Esp", values="Horas", aggfunc="sum"
)

# Workbook name, tab name, start row (row where the pivot is going to start)
pivot_table.to_excel("reports/pivot_table.xlsx", "Report", startrow=4)
