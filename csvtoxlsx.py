import pandas as pd

# Ler arquivo CSV
df = pd.read_csv('/Users/ivanmoria/Documents/GitHub/sheet_view_extract_bibref/metricas.csv')

# Salvar como Excel
df.to_excel('/Users/ivanmoria/Documents/GitHub/sheet_view_extract_bibref/metricas.xlsx', index=False)
