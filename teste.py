import pandas as pd
import numpy as np
import mca  

# Ler apenas as colunas Public e Design
counts = pd.read_csv(
    '/Users/ivanmoria/Desktop/17-16.csv',
    usecols=['Public', 'Design']
)

# Contar combinações repetidas e ordenar em ordem decrescente
grouped_counts = counts.groupby(['Public', 'Design']).size().reset_index(name='Count')
grouped_counts = grouped_counts.sort_values(by='Count', ascending=False)

# Mostrar resultado
print(grouped_counts)
