import pandas as pd
import os

# Ler arquivo CSV
df = pd.read_csv('/Users/ivanmoria/Desktop/public_counts_grouped.csv')

# Salvar como Excel
df.to_excel('/Users/ivanmoria/Desktop/publicd.xlsx', index=False)





# Caminho do arquivo Excel
xlsx_path = '/Users/ivanmoria/Desktop/aqui.xlsx'

# Nome do arquivo CSV (mesmo nome do Excel, com extensão .csv)
csv_path = os.path.splitext(xlsx_path)[0] + '.csv'

# Ler o arquivo Excel (primeira planilha por padrão)
df = pd.read_excel(xlsx_path)

# Salvar como CSV
df.to_csv(csv_path, index=False)

print(f"✔️ Arquivo CSV salvo em: {csv_path}")
