import pandas as pd
import re

file_path = '17.xlsx'
output_path = '17_processado.xlsx'

def extrair_e_organizar_completo():
    # 1. Carregar o arquivo
    df = pd.read_excel(file_path)
    novas_linhas = []

    for index, row in df.iterrows():
        dados_linha = row.to_dict()

        # --- SEPARAÇÃO DAS REFERÊNCIAS ---
        # Procura a coluna de referência (ajusta para 'Ref' ou 'ref' se necessário)
        col_ref_text = 'Ref' if 'Ref' in df.columns else 'ref'
        texto_refs = str(row[col_ref_text]) if col_ref_text in row and pd.notna(row[col_ref_text]) else ""
        
        # Procura a contagem original de referências
        col_num_ref = 'Num de Ref' if 'Num de Ref' in df.columns else 'num de ref'
        qtd_esperada_ref = int(row[col_num_ref]) if col_num_ref in row and pd.notna(row[col_num_ref]) else 0
        
        if texto_refs:
            partes = [p.strip() for p in re.split(r'\n\s*\n', texto_refs) if p.strip()]
            num_colunas_refs = max(qtd_esperada_ref, len(partes))
            for i in range(1, num_colunas_refs + 1):
                dados_linha[f'ref{i}'] = partes[i-1] if (i-1) < len(partes) else ""

        # --- SEPARAÇÃO E CONTAGEM DOS AUTORES ---
        col_autores_orig = 'Autores' if 'Autores' in df.columns else 'autores'
        texto_autores = str(row[col_autores_orig]) if col_autores_orig in row and pd.notna(row[col_autores_orig]) else ""
        
        if texto_autores:
            lista_autores = [a.strip() for a in texto_autores.split(',') if a.strip()]
            dados_linha['num de autor'] = len(lista_autores) # Nova coluna de contagem
            for i, autor in enumerate(lista_autores, start=1):
                dados_linha[f'autor{i}'] = autor
        else:
            dados_linha['num de autor'] = 0

        novas_linhas.append(dados_linha)

    # 2. Criar o DataFrame com todas as colunas
    df_final = pd.DataFrame(novas_linhas)

    # --- ORGANIZAÇÃO DA ORDEM DAS COLUNAS ---
    
    # Definimos a ordem do cabeçalho (mantendo as originais que mencionou)
    ordem_cabecalho = [
        'Num', 'Autores', 'Titulo', 'Public', 'Design', 'Afiliation', 
        'Region', 'country', 'Abstract', 'Num de Ref', 'num de autor'
    ]
    
    # 1. Filtrar colunas do cabeçalho que existem (tratando maiúsculas/minúsculas)
    colunas_ordenadas = []
    for col_alvo in ordem_cabecalho:
        for col_real in df_final.columns:
            if col_real.lower() == col_alvo.lower():
                colunas_ordenadas.append(col_real)
                break
    
    # 2. Adicionar os autores separados (autor1, autor2, ...)
    cols_autor_dinamicas = sorted([c for c in df_final.columns if c.startswith('autor') and c[5:].isdigit()], 
                                  key=lambda x: int(re.search(r'\d+', x).group()))
    colunas_ordenadas.extend(cols_autor_dinamicas)
    
    # 3. Adicionar as referências separadas (ref1, ref2, ...)
    cols_ref_dinamicas = sorted([c for c in df_final.columns if c.startswith('ref') and c[3:].isdigit()], 
                                key=lambda x: int(re.search(r'\d+', x).group()))
    colunas_ordenadas.extend(cols_ref_dinamicas)
    
    # 4. Adicionar qualquer outra coluna que possa ter sobrado
    for col in df_final.columns:
        if col not in colunas_ordenadas:
            colunas_ordenadas.append(col)

    # Aplicar a ordem e salvar
    df_final = df_final[colunas_ordenadas]
    df_final.to_excel(output_path, index=False)
    
    print(f"Processamento concluído! As colunas originais foram mantidas e organizadas.")

if __name__ == "__main__":
    extrair_e_organizar_completo()