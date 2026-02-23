
import os
import pandas as pd
import re
csv_file = "/Users/ivanmoria/Desktop/PYMT/exported_table.csv"

output_folder = "/Users/ivanmoria/Desktop/bibs"


os.makedirs(output_folder, exist_ok=True)

# Lê a tabela CSV
df = pd.read_csv(csv_file, sep=',')

# Divide a coluna 'Ref' em referências separadas (assume duplo enter entre elas)
all_refs = df['Ref'].dropna().str.cat(sep='\n\n').split('\n\n')
all_refs = [ref.strip() for ref in all_refs if ref.strip()]

# Contadores
ref_index = 0
file_counter = 1

# Itera sobre a coluna "Num de Ref" para criar os arquivos .bib
for _, row in df.iterrows():
    try:
        num_refs = int(row['Num de Ref'])
    except (ValueError, KeyError):
        print(f"Linha inválida, pulando...")
        continue

    selected_refs = all_refs[ref_index:ref_index + num_refs] if num_refs > 0 else []

    output_path = os.path.join(output_folder, f"{file_counter}.bib")
    with open(output_path, 'w', encoding='utf-8') as f_out:
        if selected_refs:
            f_out.write('\n\n'.join(selected_refs))
        else:
            f_out.write('')  # Cria arquivo vazio

    print(f"{file_counter}.bib criado com {len(selected_refs)} referências.")
    ref_index += num_refs
    file_counter += 1

print("Todos os arquivos .bib foram criados com sucesso.")
