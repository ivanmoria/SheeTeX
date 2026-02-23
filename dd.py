import pandas as pd
import re
import mca
import matplotlib.pyplot as plt
import os
import matplotlib.pyplot as plt
from adjustText import adjust_text  # pip install adjustText
import numpy as np


def plot_mca_structural(coords_rows, coords_cols, binary_df, title, top_n_vars=20):
    plt.figure(figsize=(12, 10))

    # Limites expandidos para evitar agrupamento visual
    x_min = min(coords_rows[:, 0].min(), coords_cols[:, 0].min()) - 0.5
    x_max = max(coords_rows[:, 0].max(), coords_cols[:, 0].max()) + 0.5
    y_min = min(coords_rows[:, 1].min(), coords_cols[:, 1].min()) - 0.5
    y_max = max(coords_rows[:, 1].max(), coords_cols[:, 1].max()) + 0.5

    plt.xlim(x_min, x_max)
    plt.ylim(y_min, y_max)

    # Plotar observações (linhas)
    x_rows = coords_rows[:, 0]
    y_rows = coords_rows[:, 1]
    plt.scatter(x_rows, y_rows, c='blue', label='Observações', alpha=0.6)

    # Frequência das variáveis para escolher as mais importantes
    freq = binary_df.sum().sort_values(ascending=False)
    top_vars = freq.head(top_n_vars).index.tolist()

    # Plotar variáveis (colunas)
    x_cols = coords_cols[:, 0]
    y_cols = coords_cols[:, 1]
    plt.scatter(x_cols, y_cols, c='red', label='Variáveis', marker='s', alpha=0.8)

    # Ajustar textos para linhas (todas as observações)
    texts = []
    for i, label in enumerate(binary_df.index):
        texts.append(plt.text(x_rows[i], y_rows[i], str(label), fontsize=9, color='blue'))

    # Ajustar textos para colunas (só top variáveis)
    for i, var in enumerate(binary_df.columns):
        if var in top_vars:
            texts.append(plt.text(x_cols[i], y_cols[i], var, fontsize=10, color='red'))

    # Ajustar textos para evitar sobreposição
    adjust_text(texts,
                arrowprops=dict(arrowstyle='-', color='gray', lw=0.5),
                expand_points=(2, 2),
                expand_text=(1.5, 1.5),
                force_points=0.3,
                force_text=0.3,
                max_iter=2000)

    plt.title(f"Mapa Conceitual MCA - {title}")
    plt.xlabel("Dimensão 1")
    plt.ylabel("Dimensão 2")
    plt.axhline(0, color='gray', lw=0.5)
    plt.axvline(0, color='gray', lw=0.5)
    plt.legend()
    plt.grid(True)
    plt.tight_layout()
    plt.show()



# === Arquivo de entrada ===
CSV_PATH = "/Users/ivanmoria/Desktop/17-7.csv"  # Substitua com o caminho correto do seu arquivo

# === Mapeamentos ===
GROUP_MAP_PUBLIC = {
    'children': 'Children', 'child': 'Children',
    'autism': 'Autism', 'asd': 'Autism',
    'professionals': 'Professionals', 'educators': 'Professionals',
    'dementia': 'Dementia',
    'covid-19': 'COVID-19',
    'young people': 'Young people', 'adolescents': 'Young people',
    'music therapy community': 'Music Therapy Community', 'students': 'Music Therapy Community',
    'parkinson`s disease': 'Parkinson`s Disease',
    'premature infants': 'Premature Infants', 'prenatal': 'Premature Infants',
    'nicu': 'Premature Infants', 'early childhood': 'Premature Infants',
    'older persons': 'Elderly', 'elderly population': 'Elderly',
    'end-of-life care': 'Elderly', 'elderly': 'Elderly',
    'mental health': 'Mental Health', 'schizophrenia': 'Mental Health',
    'performance anxiety': 'Mental Health', 'depression': 'Mental Health',
    'psychiatric': 'Mental Health', 'schizophrenic': 'Mental Health',
    'bipolar disorders': 'Mental Health', 'anxiety': 'Mental Health',
    'psychiatric group music therapy': 'Mental Health',
    'burnout': 'Mental Health',
    'mental illness in college music students': 'Mental Health',
    'theoretical study': 'Theoretical Study',
}

GROUP_MAP_DESIGN = {
    'theoretical study': 'Theoretical Study',
    'theoretical research': 'Theoretical Study',
}

# === Funções auxiliares ===
def normalize(text):
    if pd.isna(text):
        return ''
    text = str(text).lower().strip()
    return re.sub(r'\s+', ' ', text)

def map_group(term, group_map):
    return group_map.get(term, term)

def split_and_map(cell, group_map):
    if pd.isna(cell) or not str(cell).strip():
        return []
    parts = [normalize(p) for p in str(cell).replace('\n', ' ').replace('\r', ' ').split(',')]
    return [map_group(p, group_map) for p in parts if p]

def prepare_binary_matrix(df, column, group_map):
    binary_rows = []
    for _, row in df.iterrows():
        items = split_and_map(row[column], group_map)
        binary_rows.append({item: 1 for item in items})
    binary_df = pd.DataFrame(binary_rows).fillna(0)
    return binary_df

def run_mca(binary_df):
    mca_ben = mca.MCA(binary_df, benzecri=True)
    coords_rows = mca_ben.fs_r(N=2)  # coordinates for rows
    coords_cols = mca_ben.fs_c(N=2)  # coordinates for columns
    return mca_ben, coords_rows, coords_cols

def generate_report(df, binary_df, mca_obj, coords_rows, coords_cols, column_name):
    report = {}

    # Frequências
    freq = binary_df.sum().sort_values(ascending=False)
    report['frequencies'] = freq

    # Total de respostas
    report['total_responses'] = len(binary_df)

    # MCA - Autovalores e variância explicada
    report['eigenvalues'] = mca_obj.L
    report['variance_explained'] = mca_obj.expl_var

    # Coordenadas principais (Dim1 e Dim2) para as variáveis
    coords_df = pd.DataFrame(coords_cols, index=binary_df.columns, columns=['Dim1', 'Dim2'])
    report['coordinates'] = coords_df

    # Preparar DataFrame para salvar CSV
    freq_df = freq.reset_index()
    freq_df.columns = [column_name + '_Group', 'Frequency']

    # Coordenadas MCA das variáveis
    coords_df_reset = coords_df.reset_index()
    coords_df_reset.columns = [column_name + '_Group', 'Dim1', 'Dim2']

    # Unir frequências e coordenadas
    full_report = freq_df.merge(coords_df_reset, on=column_name + '_Group', how='left')

    return report, full_report

def save_report_csv(df_report, filename):
    desktop_path = os.path.expanduser("~/Desktop")
    full_path = os.path.join(desktop_path, filename)
    df_report.to_csv(full_path, index=False)
    print(f"Relatório salvo em: {full_path}")

def plot_mca(coords_rows, coords_cols, binary_df, title):
    plt.figure(figsize=(8, 6))
    for i, label in enumerate(binary_df.index):
        plt.scatter(coords_rows[i, 0], coords_rows[i, 1], color='blue', alpha=0.5)
        plt.text(coords_rows[i, 0], coords_rows[i, 1], str(label), fontsize=8)
    for i, var in enumerate(binary_df.columns):
        var_coord = coords_cols[i]
        plt.text(var_coord[0], var_coord[1], var, fontsize=10, color='red')
    plt.title(f"MCA - {title}")
    plt.xlabel("Dimension 1")
    plt.ylabel("Dimension 2")
    plt.grid(True)
    plt.tight_layout()
    plt.show()

# === Execução principal ===
df = pd.read_csv(CSV_PATH)

# Análise Public
binary_public = prepare_binary_matrix(df, 'Public', GROUP_MAP_PUBLIC)
mca_public, coords_rows_public, coords_cols_public = run_mca(binary_public)
report_public, full_report_public = generate_report(df, binary_public, mca_public, coords_rows_public, coords_cols_public, 'Public')

# Análise Design
binary_design = prepare_binary_matrix(df, 'Design', GROUP_MAP_DESIGN)
mca_design, coords_rows_design, coords_cols_design = run_mca(binary_design)
report_design, full_report_design = generate_report(df, binary_design, mca_design, coords_rows_design, coords_cols_design, 'Design')

# Juntar os dois relatórios completos em um só CSV com duas abas (se quiser, pode usar Excel)
with pd.ExcelWriter(os.path.expanduser("~/Desktop/relatorio_analise.xlsx")) as writer:
    full_report_public.to_excel(writer, sheet_name='Public', index=False)
    full_report_design.to_excel(writer, sheet_name='Design', index=False)

print("Resumo Public:")
print(report_public['frequencies'])
print("\nResumo Design:")
print(report_design['frequencies'])

print(f"\nTotal de respostas analisadas em Public: {report_public['total_responses']}")
print(f"Total de respostas analisadas em Design: {report_design['total_responses']}")

print("\nEigenvalues (Public):", report_public['eigenvalues'])
print("Variance Explained (Public):", report_public['variance_explained'])

print("\nEigenvalues (Design):", report_design['eigenvalues'])
print("Variance Explained (Design):", report_design['variance_explained'])

# Plot MCA para visualização
#plot_mca(coords_rows_public, coords_cols_public, binary_public, 'Public')
#plot_mca(coords_rows_design, coords_cols_design, binary_design, 'Design')
plot_mca_structural(coords_rows_public, coords_cols_public, binary_public, "Público")
plot_mca_structural(coords_rows_design, coords_cols_design, binary_design, "Design")
