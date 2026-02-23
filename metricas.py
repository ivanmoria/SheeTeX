from collections import Counter, defaultdict
import pandas as pd
import re
import os
from PyQt5.QtCore import QStandardPaths


# Mapeamento para normalizar nomes de países
mapa_paises = {
    "Brazi": "Brazil",
    "U.S": "USA",
    "U.S.A": "USA",
    "United States": "USA",
    "United States of America": "USA",
    "South Korea": "Republic of Korea",
    "Korea": "Republic of Korea",
    "UK": "United Kingdom",
    "United Kingdom": "United Kingdom",
    "Canada/Bermuda": "Canada",
    "Argentina and Spain": "Argentina",
    "Argentina and USA": "Argentina",
    "USA and South Korea": "USA",
    "USA and Taiwan": "USA",
    "USA/Puerto Rico": "USA",
}

def normalizar_paises(pais):
    if pd.isna(pais):
        return []

    pais = pais.strip()

    # Expressão regular para dividir corretamente os países
    pattern = re.compile(
        r'((?:St\s+[^\s,;/]+(?:\s+[^\s,;/]+)*)|(?:[^,;/&]+(?:\s*&\s*[^,;/&]+)+)|[^,;/&]+)', 
        re.IGNORECASE
    )

    matches = pattern.findall(pais)
    paises_normalizados = []

    for match in matches:
        p = match.strip()
        if not p:
            continue
        p = re.sub(r'\s*&\s*', ' & ', p)  # Padroniza espaços ao redor de &
        p = mapa_paises.get(p, p)  # Normaliza nome se estiver no dicionário
        paises_normalizados.append(p)

    return paises_normalizados

def calcular_metricas(dataframe: pd.DataFrame) -> str:
    """Calcula métricas descritivas e contagens únicas por país e região."""

    if dataframe.empty:
        return "Nenhum dado carregado."

    # Totais principais
    total_refs = dataframe["Num"].nunique() if "Num" in dataframe.columns else len(dataframe)
    total_autores = dataframe["Autores"].nunique() if "Autores" in dataframe.columns else 0
    total_paises = dataframe["country"].nunique() if "country" in dataframe.columns else 0
    total_regioes = dataframe["Region"].nunique() if "Region" in dataframe.columns else 0
    total_anos = dataframe["Ano"].nunique() if "Ano" in dataframe.columns else 0

    # 📌 Autores com 2+ aparições
    autores_frequentes = []
    if "Autores" in dataframe.columns:
        lista_autores = dataframe["Autores"].dropna().apply(lambda x: [a.strip() for a in x.split(",")])
        todos_autores = [autor for sublist in lista_autores for autor in sublist if autor]
        contagem_autores = Counter(todos_autores)
        autores_frequentes = [(a, f) for a, f in contagem_autores.items()]
    #    autores_frequentes = [(a, f) for a, f in contagem_autores.items() if f >= 2]

    # 📌 Contagem por país (único por Num)
    contagem_paises = defaultdict(int)
    if {"Num", "country"}.issubset(dataframe.columns):
        df_pais = dataframe[["Num", "country"]].dropna().drop_duplicates()
        df_pais = df_pais[df_pais["country"].str.strip() != ""]
        for _, row in df_pais.iterrows():
            for pais in set(normalizar_paises(row["country"])):
                contagem_paises[pais] += 1

    # 📌 Contagem por região (único por Num)
    contagem_regioes = []
    if {"Num", "Region"}.issubset(dataframe.columns):
        df_regiao = dataframe[["Num", "Region"]].dropna().drop_duplicates()
        df_regiao = df_regiao[df_regiao["Region"].str.strip() != ""]
        contagem_regioes = df_regiao.groupby("Region")["Num"].nunique().items()

    # 🔎 Montagem do relatório
    texto = (
        f"📊 {total_refs} referências únicas (Num) | "
        f"{total_autores} autores únicos | "
        f"{total_paises} países | "
        f"{total_regioes} regiões | "
        f"{total_anos} anos distintos\n"
    )

    texto += "\n👤 Autores:\n" if autores_frequentes else "\n👤 Nenhum autor.\n"
    for autor, freq in sorted(autores_frequentes, key=lambda x: x[1], reverse=True):
        texto += f" - {autor}: {freq}\n"

    if contagem_paises:
        texto += "\n🌍 Contagem de trabalhos únicos por país:\n"
        for pais, freq in sorted(contagem_paises.items(), key=lambda x: x[1], reverse=True):
            texto += f" - {pais}: {freq}\n"
    else:
        texto += "\n🌍 Nenhum dado de país disponível.\n"

    if contagem_regioes:
        texto += "\n🗺️ Contagem de trabalhos únicos por região:\n"
        regioes_ordenadas = sorted(contagem_regioes, key=lambda x: x[1], reverse=True)
        for regiao, freq in regioes_ordenadas:
            texto += f" - {regiao}: {freq}\n"
    else:
        texto += "\n🗺️ Nenhum dado de região disponível.\n"

    return texto


def exportar_metricas_texto_para_csv(texto: str):
        """Recebe o texto das métricas formatado e salva como CSV na pasta principal."""

        blocos = {
            "🗺️ Contagem de trabalhos únicos por região:": [],
            "🌍 Contagem de trabalhos únicos por país:": [],
            "👤 Autores:": []  # Corrigido para corresponder ao texto gerado
        }

        bloco_atual = None
        for linha in texto.strip().splitlines():
            linha = linha.strip()
            if linha in blocos:
                bloco_atual = linha
            elif linha and bloco_atual and linha != "👤 Nenhum autor.":
                blocos[bloco_atual].append(linha)

        def separar_nome_valor(lista):
            nomes, valores = [], []
            for item in lista:
                if ':' in item:
                    nome, valor = item.split(':', 1)
                    nomes.append(nome.strip(' -'))
                    valores.append(valor.strip())
                else:
                    nomes.append(item.strip(' -'))
                    valores.append("")
            return nomes, valores

        # Corrigido para usar as chaves corretas
        regioes, qtd_regioes = separar_nome_valor(blocos["🗺️ Contagem de trabalhos únicos por região:"])
        paises, qtd_paises = separar_nome_valor(blocos["🌍 Contagem de trabalhos únicos por país:"])
        autores, qtd_autores = separar_nome_valor(blocos["👤 Autores:"])

        max_linhas = max(len(regioes), len(paises), len(autores))
        while len(regioes) < max_linhas:
            regioes.append("")
            qtd_regioes.append("")
        while len(paises) < max_linhas:
            paises.append("")
            qtd_paises.append("")
        while len(autores) < max_linhas:
            autores.append("")
            qtd_autores.append("")

        df = pd.DataFrame({
            "Região": regioes,
            "Num1": qtd_regioes,
            "País": paises,
            "Num2": qtd_paises,
            "Autor": autores,
            "Num3": qtd_autores
        })
        desktop = QStandardPaths.writableLocation(QStandardPaths.DesktopLocation)

        # Caminho para a pasta PYMT dentro da área de trabalho
        pasta_pymt = os.path.join(desktop, "PYMT")

        # Cria a pasta se ela não existir
        os.makedirs(pasta_pymt, exist_ok=True)

        # Caminho completo do arquivo a ser salvo
        file_path = os.path.join(pasta_pymt, "metricas.csv")


        try:
            df.to_csv(file_path, index=False)
            print(f"Arquivo salvo com sucesso em: {file_path}")
        except Exception as e:
            print(f"Erro ao salvar arquivo: {e}")


def exportar_metricas_texto_para_csv(texto: str):
        """Recebe o texto das métricas formatado e salva como CSV na pasta principal."""

        blocos = {
            "🗺️ Contagem de trabalhos únicos por região:": [],
            "🌍 Contagem de trabalhos únicos por país:": [],
            "👤 Autores:": []  # Corrigido para corresponder ao texto gerado
        }

        bloco_atual = None
        for linha in texto.strip().splitlines():
            linha = linha.strip()
            if linha in blocos:
                bloco_atual = linha
            elif linha and bloco_atual and linha != "👤 Nenhum autor.":
                blocos[bloco_atual].append(linha)

        def separar_nome_valor(lista):
            nomes, valores = [], []
            for item in lista:
                if ':' in item:
                    nome, valor = item.split(':', 1)
                    nomes.append(nome.strip(' -'))
                    valores.append(valor.strip())
                else:
                    nomes.append(item.strip(' -'))
                    valores.append("")
            return nomes, valores

        # Corrigido para usar as chaves corretas
        regioes, qtd_regioes = separar_nome_valor(blocos["🗺️ Contagem de trabalhos únicos por região:"])
        paises, qtd_paises = separar_nome_valor(blocos["🌍 Contagem de trabalhos únicos por país:"])
        autores, qtd_autores = separar_nome_valor(blocos["👤 Autores:"])

        max_linhas = max(len(regioes), len(paises), len(autores))
        while len(regioes) < max_linhas:
            regioes.append("")
            qtd_regioes.append("")
        while len(paises) < max_linhas:
            paises.append("")
            qtd_paises.append("")
        while len(autores) < max_linhas:
            autores.append("")
            qtd_autores.append("")

        df = pd.DataFrame({
            "Região": regioes,
            "Num1": qtd_regioes,
            "País": paises,
            "Num2": qtd_paises,
            "Autor": autores,
            "Num3": qtd_autores
        })
        desktop = QStandardPaths.writableLocation(QStandardPaths.DesktopLocation)

        # Caminho para a pasta PYMT dentro da área de trabalho
        pasta_pymt = os.path.join(desktop, "PYMT")

        # Cria a pasta se ela não existir
        os.makedirs(pasta_pymt, exist_ok=True)

        # Caminho completo do arquivo a ser salvo
        file_path = os.path.join(pasta_pymt, "metricas.csv")


        try:
            df.to_csv(file_path, index=False)
            print(f"Arquivo salvo com sucesso em: {file_path}")
        except Exception as e:
            print(f"Erro ao salvar arquivo: {e}")
            
def exportar_metricas_detalhe_para_csv(texto: str):
        """Recebe o texto das métricas formatado e salva como CSV na pasta principal."""

        blocos = {
            "🗺️ Contagem de trabalhos únicos por região:": [],
            "🌍 Contagem de trabalhos únicos por país:": [],
            "👤 Autores:": []  # Corrigido para corresponder ao texto gerado
        }

        bloco_atual = None
        for linha in texto.strip().splitlines():
            linha = linha.strip()
            if linha in blocos:
                bloco_atual = linha
            elif linha and bloco_atual and linha != "👤 Nenhum autor.":
                blocos[bloco_atual].append(linha)

        def separar_nome_valor(lista):
            nomes, valores = [], []
            for item in lista:
                if ':' in item:
                    nome, valor = item.split(':', 1)
                    nomes.append(nome.strip(' -'))
                    valores.append(valor.strip())
                else:
                    nomes.append(item.strip(' -'))
                    valores.append("")
            return nomes, valores

        # Corrigido para usar as chaves corretas
        regioes, qtd_regioes = separar_nome_valor(blocos["🗺️ Contagem de trabalhos únicos por região:"])
        paises, qtd_paises = separar_nome_valor(blocos["🌍 Contagem de trabalhos únicos por país:"])
        autores, qtd_autores = separar_nome_valor(blocos["👤 Autores:"])

        max_linhas = max(len(regioes), len(paises), len(autores))
        while len(regioes) < max_linhas:
            regioes.append("")
            qtd_regioes.append("")
        while len(paises) < max_linhas:
            paises.append("")
            qtd_paises.append("")
        while len(autores) < max_linhas:
            autores.append("")
            qtd_autores.append("")

        df = pd.DataFrame({
            "Região": regioes,
            "Num1": qtd_regioes,
            "País": paises,
            "Num2": qtd_paises,
            "Autor": autores,
            "Num3": qtd_autores
        })
        desktop = QStandardPaths.writableLocation(QStandardPaths.DesktopLocation)

        # Caminho para a pasta PYMT dentro da área de trabalho
        pasta_pymt = os.path.join(desktop, "PYMT")

        # Cria a pasta se ela não existir
        os.makedirs(pasta_pymt, exist_ok=True)

        # Caminho completo do arquivo a ser salvo
        file_path = os.path.join(pasta_pymt, "metricas_detalhe.csv")


        try:
            df.to_csv(file_path, index=False)
            print(f"Arquivo salvo com sucesso em: {file_path}")
        except Exception as e:
            print(f"Erro ao salvar arquivo: {e}")



def calcular_metricas_detalhadas(dataframe: pd.DataFrame) -> str:
    if dataframe.empty:
        return "⚠️ Nenhum dado carregado."

    texto = "📌 MÉTRICAS DETALHADAS\n"

    def contar_percentual(series, titulo, emoji, n_total=None):
        series = series.dropna()
        contagem = series.value_counts()
        total = n_total if n_total else len(series)

        bloco = f"\n{emoji} {titulo}:\n"
        for valor, count in contagem.items():
            perc = 100 * count / total
            bloco += f" - {valor}: {count} ({perc:.1f}%)\n"
        return bloco

    def expandir_multiplos(valores):
        return [item.strip() for v in valores.dropna() for item in v.split(",") if item.strip()]

    def normalizar_paises(texto):
        texto = texto.strip().lower()
        paises = [p.strip().capitalize() for p in texto.split(",") if p]
        return paises

    def agrupar_termos(lista):
        mapa_equivalencias = {
         
            "Autism":  "ASD",
            "autism": "ASD",
            "autistic": "ASD",
            "children": "Children",
            "child": "Children",
            "adolescents": "Adolescents",
            "adolescent": "Adolescents",
            "student": "Students",
            "students": "Students",
            "GIM": "Guided Imagery and Music",
            "Burnout": "Burnout",
            "telehealth": "Telehealth",
            "developmental disabilities": "Developmental Disability",
            "Experience report": "Experience Report",
             "experience report": "Experience Report",
  "Theoretical study, Quali": "Theoretical study, Qualitative",
              "survey": "Survey",
                  "Quali": "Qualitative",
                    "Theoretical Study, quali": "Qualitative",
                                   "Parkinson`s Disease": " Parkinson's Disease",

                                    "Theoretical study, clinical practice": "Clinical Practice",
                  "Experiênce Report   ": "Experience Report",
                                               "Opinion, Survey": "Survey",
                      "opinion": "Opinion",
  " Opinion, Survey": "Survey",
  "Quali-Quantitative": "Quali-Quanti",
    "Quali- Quantitative": "Quali-Quanti",
        "Theoretical study, Quanti": "Quantitative",
                "Theoretical Study, Opinion": "Opinion",
 "Theoretical study, Quali": "Theoretical study, Qualitative",
            "elderly": "Elderly",
            "older adults": "Elderly",
            "Parkinson`s Disease": "Parkinson's Disease",
            "mental health": "Mental Health",
            "developmental disability": "Developmental Disability"
            # Adicione outros agrupamentos desejados aqui
        }

        normalizados = []
        for termo in lista:
            chave = termo.strip().lower()
            termo_agrupado = mapa_equivalencias.get(chave, termo.strip())
            normalizados.append(termo_agrupado)
        return normalizados

    # Tipos de publicação
    if "Public" in dataframe.columns:
        df_public = dataframe[["Num", "Public"]].dropna()
        df_public_unicos = df_public.drop_duplicates(subset=["Num", "Public"])

        termos_expandidos = expandir_multiplos(df_public_unicos["Public"])
        termos_agrupados = agrupar_termos(termos_expandidos)

        contagem = Counter(termos_agrupados)
        total = len(df_public_unicos["Num"].unique())

        texto += "\n📚 Tipos de publicação mais frequentes (contagem única por Num):\n"
        for termo, count in contagem.most_common(20):
            perc = 100 * count / total
            texto += f" - {termo}: {count} ({perc:.1f}%)\n"
    else:
        texto += "\n📚 Coluna 'Public' não encontrada.\n"

    # Tipos de delineamento
    if "Design" in dataframe.columns:
        texto += contar_percentual(dataframe["Design"], "Tipos de delineamento mais frequentes", "🧪")
    else:
        texto += "\n🧪 Coluna 'Design' não encontrada.\n"

    # Países
    if "country" in dataframe.columns:
        paises_normalizados = dataframe["country"].dropna().apply(normalizar_paises)
        todos_paises = [pais for lista in paises_normalizados for pais in lista]
        contagem_paises = Counter(todos_paises)
        total_refs = len(dataframe["Num"].unique()) if "Num" in dataframe.columns else len(dataframe)

        texto += "\n🌍 Países mais frequentes:\n"
        for pais, count in contagem_paises.most_common(10):
            perc = 100 * count / total_refs
            texto += f" - {pais}: {count} referências ({perc:.1f}%)\n"
    else:
        texto += "\n🌍 Coluna 'country' não encontrada.\n"

    # Regiões
    if "Region" in dataframe.columns:
        texto += contar_percentual(dataframe["Region"], "Regiões mais frequentes", "🗺️")
    else:
        texto += "\n🗺️ Coluna 'Region' não encontrada.\n"

    # Autores
    if "Autores" in dataframe.columns:
        lista_autores = dataframe["Autores"].dropna().apply(lambda x: [a.strip() for a in x.split(",")])
        todos_autores = [autor for sublist in lista_autores for autor in sublist if autor]
        contagem_autores = Counter(todos_autores)
        num_refs = len(lista_autores)
        total_autores = sum(len(sublist) for sublist in lista_autores)
        media_autores = total_autores / num_refs if num_refs > 0 else 0

        texto += "\n👤 Autores mais frequentes:\n"
        for autor, count in contagem_autores.most_common(10):
            texto += f" - {autor}: {count} aparições\n"

        texto += f"\n👥 Média de autores por referência: {media_autores:.2f}\n"
    else:
        texto += "\n👤 Coluna 'Autores' não encontrada.\n"

    # Anos
    if "Ano" in dataframe.columns:
        contagem_anos = dataframe["Ano"].dropna().value_counts().sort_index()
        texto += "\n📅 Distribuição por ano:\n"
        for ano, count in contagem_anos.items():
            texto += f" - {ano}: {count} referência(s)\n"
    else:
        texto += "\n📅 Coluna 'Ano' não encontrada.\n"

    # Combinação país-autor
    if "country" in dataframe.columns and "Autores" in dataframe.columns:
        texto += "\n🌐 Combinações país-autor mais frequentes:\n"
        combinacoes = []
        for _, row in dataframe.dropna(subset=["country", "Autores"]).iterrows():
            paises = normalizar_paises(row["country"])
            autores = [a.strip() for a in row["Autores"].split(",") if a.strip()]
            combinacoes.extend([(pais, autor) for pais in paises for autor in autores])

        contagem_combinacoes = Counter(combinacoes)
        for (pais, autor), count in contagem_combinacoes.most_common(10):
            texto += f" - {autor} ({pais}): {count} vezes\n"

    return texto
