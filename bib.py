import pandas as pd
import json
import re
from processador_referencias import extrair_campos_apa, extrair_autores_completos

def expandir_linhas_por_quebra(df: pd.DataFrame, coluna='Ref', separador=None) -> pd.DataFrame:
    """
    Divide o conteúdo da coluna pelo separador e cria uma linha para cada parte,
    replicando os dados das outras colunas.
    """
    if separador is None:
        raise ValueError("Você deve informar o separador correto")

    linhas_expandidas = []
    for idx, row in df.iterrows():
        conteudo = str(row[coluna])
        partes = conteudo.split(separador)
        for parte in partes:
            nova_linha = row.copy()
            nova_linha[coluna] = parte.strip()
            linhas_expandidas.append(nova_linha)
    df_exp = pd.DataFrame(linhas_expandidas).reset_index(drop=True)
    return df_exp
def processar_entrada_bibtex(df: pd.DataFrame) -> str:
    """
    Processa um DataFrame com uma coluna 'Ref' contendo referências no estilo APA.
    Extrai campos e converte para formato BibTeX.
    """
    try:
        # Extrai todos os textos da coluna 'Ref' como uma única string
        textos = df['Ref'].dropna().astype(str).tolist()
        entrada = "\n\n".join(textos).strip()

        if not entrada:
            return ""

        # Separa as referências por duas quebras de linha
        referencias_separadas = re.split(r'\n{2,}', entrada)

        bibtex_total = ""
        for idx, ref in enumerate(referencias_separadas, 1):
            campos = extrair_campos_apa(ref)

            bibtex = f"@article{{ref{idx},\n"
            for campo, valor in campos.items():
                if valor:
                    bibtex += f"  {campo} = {{{valor}}},\n"
            bibtex += f"  note = {{{ref.strip()}}}\n}}\n\n"

            bibtex_total += bibtex

        return bibtex_total

    except Exception as e:
        return f"Erro ao processar entrada bibtex: {e}"
def processar_visualizacao_formatada(df: pd.DataFrame) -> str:

    try:
        df_exp = expandir_linhas_por_quebra(df, coluna='Ref', separador='\n\n')

        df_sel = df_exp[['Num', 'Num de Ref', 'Ref']]
        texto_formatado = json.dumps(df_sel.to_dict(orient='records'), indent=4, ensure_ascii=False)
        return texto_formatado
    except Exception as e:
        return f"Erro ao processar visualização formatada: {e}"

def processar_estatisticas_bibtex(df: pd.DataFrame) -> str:

    try:
        df_exp = expandir_linhas_por_quebra(df, coluna='Ref', separador='\\n\\n\\')
        
        total_linhas = len(df_exp)
        num_refs_unicas = df_exp['Ref'].nunique()
        num_num_unicos = df_exp['Num'].nunique()
        num_num_ref_unicos = df_exp['Num de Ref'].nunique()

        estatisticas = (
            f"Total de linhas: {total_linhas}\n"
            f"Número de referências únicas: {num_refs_unicas}\n"
            f"Número de valores únicos em 'Num': {num_num_unicos}\n"
            f"Número de valores únicos em 'Num de Ref': {num_num_ref_unicos}\n"
        )
        return estatisticas
    except Exception as e:
        return f"Erro ao processar estatísticas bibtex: {e}"
