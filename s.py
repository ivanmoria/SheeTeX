import csv

def somar_segunda_coluna(nome_arquivo):
    soma = 0
    with open(nome_arquivo, newline='', encoding='utf-8') as csvfile:
        leitor = csv.reader(csvfile)
        next(leitor)  # Pula o cabeçalho
        for linha in leitor:
            if len(linha) >= 2:
                try:
                    soma += int(linha[1])
                except ValueError:
                    print(f"Valor inválido na linha: {linha}")
    return soma

if __name__ == "__main__":
    arquivo = '/Users/ivanmoria/Desktop/design_stats.csv' # Altere se necessário
    total = somar_segunda_coluna(arquivo)
    print(f"Soma da segunda coluna: {total}")
