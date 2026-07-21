import pandas as pd

df = pd.read_excel('JUN-17-2026.xlsx', nrows=0)
headers = df.columns.tolist()
for i, h in enumerate(headers):
    print(f"{i}: {h}")
