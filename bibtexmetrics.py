from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QTableWidget, QTableWidgetItem,
    QVBoxLayout, QWidget, QLabel, QMessageBox, QTabWidget,
    QPushButton, QHBoxLayout
)
from PyQt5.QtCore import Qt
from pathlib import Path
import bibtexparser
from collections import defaultdict
import statistics
import sys
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import numpy as np
from wordcloud import WordCloud
import matplotlib.pyplot as plt
import re
from processador_referencias import extrair_campos_apa, extrair_autores_completos
from wordcloud import STOPWORDS

stopwords = set(STOPWORDS)
wc = WordCloud(stopwords=stopwords)


def analyze_bibtex():
    # Caminho do arquivo bibtex
    bibtex_path = Path.home() / "Desktop/PYMT/references.bib"
    if not bibtex_path.exists():
        raise FileNotFoundError(f"Arquivo {bibtex_path} não encontrado.")

    with open(bibtex_path, encoding="utf-8") as bibtex_file:
        bib_database = bibtexparser.load(bibtex_file)

    entries = bib_database.entries

    # Lista das publicações completas para tabela detalhada (sem N/A)
    publicacoes_completas = []
    publicacoes_por_ano = defaultdict(int)
    tipos_publicacao = defaultdict(int)
    autores_count = defaultdict(int)
    decadas_count = defaultdict(int)

    for entry in entries:
        year = entry.get("year", "N/A")
        tipo = entry.get("ENTRYTYPE", "N/A")
        bib_id = entry.get("ID", "N/A")
        authors = entry.get("author", "N/A")
        title = entry.get("title", "N/A")

        # Ignorar entradas que tenham qualquer campo "N/A"
        if "N/A" in (year, tipo, bib_id, authors, title):
            continue

        # Publicações completas (linha da tabela)
        publicacoes_completas.append([year, tipo, bib_id, authors, title])

        # Publicações por ano
        if year.isdigit():
            year_int = int(year)
            publicacoes_por_ano[year_int] += 1
            decada = (year_int // 10) * 10
            decadas_count[decada] += 1

        # Tipos de publicação
        tipos_publicacao[tipo] += 1

        # Contagem de autores
        for autor in authors.split(" and "):
            autor = autor.strip()
            if autor:
                autores_count[autor] += 1

    # Ordenar dados para exibição
    publicacoes_por_ano = sorted(publicacoes_por_ano.items())
    tipos_publicacao = sorted(tipos_publicacao.items(), key=lambda x: x[1], reverse=True)
    autores_completos = sorted(autores_count.items(), key=lambda x: x[1], reverse=True)
    decadas = sorted(decadas_count.items())

    # Estatísticas gerais simples
    total_publicacoes = len(publicacoes_completas)
    total_autores_unicos = len(autores_count)
    anos_analizados = len(publicacoes_por_ano)
    media_pub_por_ano = round(total_publicacoes / anos_analizados, 2) if anos_analizados else 0

    estatisticas_gerais = [
        ("Total de publicações", total_publicacoes),
        ("Autores únicos", total_autores_unicos),
        ("Anos analisados", anos_analizados),
        ("Média de publicações por ano", media_pub_por_ano)
    ]

    return {     "entries": entries,
        "publicacoes_completas": publicacoes_completas,
        "publicacoes_por_ano": publicacoes_por_ano,
        "tipos_publicacao": tipos_publicacao,
        "autores_completos": autores_completos,
        "decadas": decadas,
        "estatisticas_gerais": estatisticas_gerais,
    }



class BibtexViewer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Visualizador de BibTeX")
        self.resize(1000, 600)
        self.initUI()
        self.reload_data()

    def initUI(self):
        central_widget = QWidget()
        main_layout = QVBoxLayout()
        central_widget.setLayout(main_layout)
        self.setCentralWidget(central_widget)

        # Botões
        button_layout = QHBoxLayout()

        self.export_button = QPushButton("Exportar Tabela Atual")
        self.export_button.clicked.connect(self.exportar_tabela_atual)
        button_layout.addWidget(self.export_button)
        self.reload_button = QPushButton("Recarregar Dados")
        self.reload_button.clicked.connect(self.reload_data)
        button_layout.addWidget(self.reload_button)

        # Abas com tabelas e gráficos
        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)
        main_layout.addLayout(button_layout)

        # Dicionário para armazenar figuras
        self.figures = {}

    
    def create_plots_tab_with_subtabs(self, metrics):
        # Widget principal que contém os gráficos
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # Botão para exportar gráfico ativo
        export_plot_button = QPushButton("Exportar Gráfico Atual")
        layout.addWidget(export_plot_button, alignment=Qt.AlignRight)

        # Definindo o container de abas
        subtabs = QTabWidget()
        layout.addWidget(subtabs)

        # Função para criar o canvas de gráficos
        def create_canvas():
            fig = Figure(figsize=(6, 4))
            canvas = FigureCanvas(fig)
            return fig, canvas

        # 1. Gráfico: Publicações por Ano
        fig1, canvas1 = create_canvas()
        ax1 = fig1.add_subplot(111)
        anos = [int(x[0]) for x in metrics['publicacoes_por_ano']]
        quantidades = [x[1] for x in metrics['publicacoes_por_ano']]
        ax1.plot(anos, quantidades, marker='o')
        ax1.set_title("Publicações por Ano")
        ax1.set_xlabel("Ano")
        ax1.set_ylabel("Quantidade")
        ax1.grid(True)
        tab1 = QWidget()
        l1 = QVBoxLayout(tab1)
        l1.addWidget(canvas1)
        subtabs.addTab(tab1, "Por Ano")
        self.figures["Por Ano"] = fig1

        # 2. Gráfico: Top 10 Autores
        fig2, canvas2 = create_canvas()
        ax2 = fig2.add_subplot(111)
        top_autores = metrics['autores_completos'][:10]
        autores = [x[0] for x in top_autores]
        pubs = [x[1] for x in top_autores]
        y_pos = np.arange(len(autores))
        ax2.barh(y_pos, pubs, color='lightgreen')
        ax2.set_yticks(y_pos)
        ax2.set_yticklabels(autores)
        ax2.invert_yaxis()  # Inverte o eixo Y para exibir o autor com mais publicações no topo
        ax2.set_title("Top 10 Autores")
        ax2.set_xlabel("Número de Publicações")
        tab2 = QWidget()
        l2 = QVBoxLayout(tab2)
        l2.addWidget(canvas2)
        subtabs.addTab(tab2, "Autores")
        self.figures["Autores"] = fig2

        # 3. Gráfico: Publicações por Década (pizza)
        fig3, canvas3 = create_canvas()
        ax3 = fig3.add_subplot(111)
        decadas = [x[0] for x in metrics['decadas']]
        dec_qtd = [x[1] for x in metrics['decadas']]
        ax3.pie(dec_qtd, labels=decadas, autopct='%1.1f%%', startangle=140)
        ax3.set_title("Publicações por Década")
        tab3 = QWidget()
        l3 = QVBoxLayout(tab3)
        l3.addWidget(canvas3)
        subtabs.addTab(tab3, "Décadas")
        self.figures["Décadas"] = fig3

            # Lista de palavras a serem removidas
        palavras_remover = ["van", "et", "al", "Music", "music", "International","Française Musicothérapie","Revista","therapy","with","eds","association","Association","2nd","Eds","use","der","Der","Kim","Federation", "Therapy", "therapist", "for","to","an","An","as","As",
                            "Ed", "de", "Journal", "De", "a","and","of","la","Of","in","on","online","during","world","World", "the","e","by","with","o", "da", "do", "das", "dos"]

        # Converte a lista de palavras a remover para minúsculas para garantir comparação consistente
        palavras_remover = [p.lower() for p in palavras_remover]

        # 4. Nuvens de Palavras para vários campos
        campos = ['title', 'author', 'journal', 'abstract', 'keywords']  # Ajuste os campos conforme o seu bibtex
        for campo in campos:
            fig_wc, canvas_wc = create_canvas()
            ax_wc = fig_wc.add_subplot(111)

            # Extrair texto para a nuvem
            if campo == 'author':
                # Filtra autores, permitindo que mesmo iniciais sejam incluídas
                autores = [entry.get(campo, '') for entry in metrics['entries']]
                autores_filtrados = []
                for autor in autores:
                    # Separa por espaços para pegar nomes completos
                    nomes = autor.split()
                    # Mantém todos os nomes completos (não exclui iniciais ou nomes válidos)
                    autores_filtrados.append(" ".join(nomes))  # Junta todos os nomes

                palavras = " ".join(autores_filtrados)  # Junta todos os autores válidos em uma única string
            else:
                palavras = " ".join([entry.get(campo, '') for entry in metrics['entries']])

            # Remove pontuações
            palavras = re.sub(r'[^\w\s]', '', palavras)  # Remove pontuações

            # Filtra palavras de uma única letra e palavras da lista de exclusão
            palavras = " ".join([p for p in palavras.split() if len(p) > 1 and p.lower() not in palavras_remover])

            # Criação da nuvem de palavras
            if palavras.strip():  # Verifica se existe algum texto após a remoção de pontuações
                # Adiciona as palavras da lista de stopwords, incluindo as de exclusão
                stopwords = set(palavras_remover)
                wc = WordCloud(width=800, height=400, background_color='white',
                                stopwords=stopwords, colormap='viridis', max_words=100).generate(palavras)
                ax_wc.imshow(wc, interpolation='bilinear')
                ax_wc.axis("off")
                ax_wc.set_title(f"Nuvem de Palavras: {campo.capitalize()}")

            # Adicionando a nuvem ao layout
            tab_wc = QWidget()
            lwc = QVBoxLayout(tab_wc)
            lwc.addWidget(canvas_wc)
            subtabs.addTab(tab_wc, f"Nuvem: {campo.capitalize()}")
            self.figures[f"Nuvem: {campo}"] = fig_wc






        # Adiciona a guia de gráficos na interface
        self.tabs.addTab(widget, "Gráficos")

        # Conectar a função de exportação de gráficos
        export_plot_button.clicked.connect(self.export_current_plot)


    def export_current_plot(self):
        widget = self.tabs.widget(self.tabs.count() - 1)
        if widget is None:
            QMessageBox.warning(self, "Erro", "Nenhum gráfico disponível para exportar.")
            return

        subtabs = widget.findChild(QTabWidget)
        if subtabs is None:
            QMessageBox.warning(self, "Erro", "Sub-abas de gráfico não encontradas.")
            return

        current_tab_name = subtabs.tabText(subtabs.currentIndex())
        fig = self.figures.get(current_tab_name, None)
        if fig is None:
            QMessageBox.warning(self, "Erro", "Não foi possível encontrar o gráfico para exportar.")
            return

        save_folder = Path.home() / "Desktop" / "PYMT"
        save_folder.mkdir(parents=True, exist_ok=True)

        n = 1
        while True:
            file_path = save_folder / f"PYMT_plot({n}).png"
            if not file_path.exists():
                break
            n += 1

        fig.savefig(str(file_path), dpi=300)
        QMessageBox.information(self, "Sucesso", f"Gráfico salvo em:\n{file_path}")



    def exportar_tabela_atual(self):
        index = self.tabs.currentIndex()
        if index == -1:
            QMessageBox.warning(self, "Aviso", "Nenhuma aba selecionada.")
            return

        tab_title = self.tabs.tabText(index).replace(" ", "_").lower()
        file_path = Path.home() / f"Desktop/PYMT/{tab_title}.csv"

        tab_widget = self.tabs.widget(index)
        table = tab_widget.findChild(QTableWidget)
        if not table:
            QMessageBox.warning(self, "Aviso", "Tabela não encontrada.")
            return

        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                if tab_title == "publicações_completas":  # aba Publicações completas
                    desired_columns = ["Ano", "Autores", "Título"]
                    col_indices = []
                    for col in range(table.columnCount()):
                        header_text = table.horizontalHeaderItem(col).text()
                        if header_text in desired_columns:
                            col_indices.append(col)

                    # Escreve cabeçalho com coluna numeração
                    headers = ["#", *desired_columns]
                    f.write(','.join(headers) + '\n')

                    # Escreve dados — **incluindo todas as linhas, sem pular**
                    for row in range(table.rowCount()):
                        row_data = [str(row + 1)]
                        for col in col_indices:
                            item = table.item(row, col)
                            text = item.text().strip() if item else ""
                            if text == "N/A":
                                text = ""  # substitui "N/A" por vazio

                            # Colocar aspas em autores e título para CSV, e escapar aspas internas
                            header = table.horizontalHeaderItem(col).text()
                            if header in ("Autores", "Título") and text:
                                text = text.replace('"', '""')
                                text = f'"{text}"'
                            row_data.append(text)

                        f.write(','.join(row_data) + '\n')

                else:
                    # Para as outras abas salva tudo normalmente
                    headers = [table.horizontalHeaderItem(i).text() for i in range(table.columnCount())]
                    f.write(','.join(headers) + '\n')
                    for row in range(table.rowCount()):
                        row_data = []
                        for col in range(table.columnCount()):
                            item = table.item(row, col)
                            row_data.append(item.text() if item else "")
                        f.write(','.join(row_data) + '\n')

            QMessageBox.information(self, "Sucesso", f"Tabela exportada como:\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao exportar CSV:\n{e}")



    def reload_data(self):
        self.tabs.clear()
        try:
            metrics = analyze_bibtex()
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Falha ao carregar dados: {str(e)}")
            return

        if not metrics or not metrics['publicacoes_completas']:
            QMessageBox.warning(self, "Aviso", "Nenhum dado encontrado no arquivo BibTeX.")
            return

        # Filtrar linhas com "N/A" em qualquer coluna
        filtered_pub_completas = [
            [row[0], row[3], row[4]]  # seleciona só Ano, Autores e Título
            for row in metrics['publicacoes_completas']
            if all(row[i] != "N/A" for i in [0, 3, 4])
        ]

        self.create_table_tab("Publicações completas", filtered_pub_completas,
                            ["Ano", "Autores", "Título"])
        
        # Demais abas continuam iguais
        self.create_table_tab("Publicações por ano", metrics['publicacoes_por_ano'],
                            ["Ano", "Quantidade"])
        self.create_table_tab("Autores mais frequentes", metrics['autores_completos'],
                            ["Autores", "Publicações"])
        self.create_table_tab("Tipos de publicação", metrics['tipos_publicacao'],
                            ["Tipo", "Quantidade"])
        self.create_table_tab("Publicações por década", metrics['decadas'],
                            ["Década", "Quantidade"])
        self.create_table_tab("Estatísticas gerais", metrics['estatisticas_gerais'],
                            ["Métrica", "Valor"])

        self.create_plots_tab_with_subtabs(metrics)


    def create_table_tab(self, title, data, headers):
        widget = QWidget()
        layout = QVBoxLayout()
        table = QTableWidget()
        layout.addWidget(table)
        widget.setLayout(layout)

        if len(data) == 0:
            table.setRowCount(0)
            table.setColumnCount(len(headers))
            table.setHorizontalHeaderLabels(headers)
        else:
            table.setRowCount(len(data))
            table.setColumnCount(len(headers))
            table.setHorizontalHeaderLabels(headers)

            for row_idx, row_data in enumerate(data):
                for col_idx, value in enumerate(row_data):
                    item = QTableWidgetItem(str(value))
                    item.setFlags(item.flags() ^ Qt.ItemIsEditable)  # Deixa célula não editável
                    table.setItem(row_idx, col_idx, item)

        # Definir larguras específicas para a aba "Publicações completas"
        if title == "Publicações completas":
            column_widths = {
                0: 50,    # Ano
                1: 275,   # Autores
                2: 525,    # Título
            }
        elif title == "Autores mais frequentes":
            column_widths = {
                0: 500,
                1: 100,
            }
        elif title == "Estatísticas gerais":
            column_widths = {
                0: 200,
                1: 80,
            }
        else:
            column_widths = {}

        # Aplicar as larguras
        for col, width in column_widths.items():
            table.setColumnWidth(col, width)

        self.tabs.addTab(widget, title)




def main():
    app = QApplication(sys.argv)
    window = BibtexViewer()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
