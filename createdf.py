import pandas as pd

archivo_origen = "clon_info.xlsx"
df = pd.read_excel(archivo_origen)

print(df.columns)