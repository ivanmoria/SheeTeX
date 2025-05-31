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
import matplotlib.pyplot as plt
import numpy as np

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

        self.update_button = QPushButton("Atualizar Tabelas")
        self.update_button.clicked.connect(self.reload_data)
        #button_layout.addWidget(self.update_button)
        self.update_button.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;       /* Verde */
                color: white;
                border: 2px solid #388E3C;
                border-radius: 30px;
                padding: 2px 2px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #45A049;       /* Verde um pouco mais claro ao passar o mouse */
                border-color: #2E7D32;
            }
        """)
        self.export_button = QPushButton("Exportar Tabela Atual")
        self.export_button.clicked.connect(self.exportar_tabela_atual)
        button_layout.addWidget(self.export_button)

       

        # Abas com tabelas
        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)
        main_layout.addLayout(button_layout)
        self.reload_data()
    def create_plots_tab_with_subtabs(self, metrics):
        widget = QWidget()
        layout = QVBoxLayout()
        widget.setLayout(layout)

        # Botão para exportar gráfico ativo
        export_plot_button = QPushButton("Exportar Gráfico Atual")
        layout.addWidget(export_plot_button, alignment=Qt.AlignRight)

        subtabs = QTabWidget()
        layout.addWidget(subtabs)

        def create_canvas():
            fig = Figure(figsize=(6,4))
            canvas = FigureCanvas(fig)
            return fig, canvas

        self.figures = {}

        # Gráfico 1: Publicações por Ano
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

        # Gráfico 3: Autores mais frequentes (top 10)
        fig3, canvas3 = create_canvas()
        ax3 = fig3.add_subplot(111)
        top_autores = metrics['autores_completos'][:10]
        autores = [x[0] for x in top_autores]
        pubs = [x[1] for x in top_autores]
        y_pos = np.arange(len(autores))
        ax3.barh(y_pos, pubs, color='lightgreen')
        ax3.set_yticks(y_pos)
        ax3.set_yticklabels(autores)
        ax3.invert_yaxis()
        ax3.set_title("Top 10 Autores")
        ax3.set_xlabel("Número de Publicações")
        tab3 = QWidget()
        l3 = QVBoxLayout(tab3)
        l3.addWidget(canvas3)
        subtabs.addTab(tab3, "Autores")
        self.figures["Autores"] = fig3

        # Gráfico 4: Publicações por Década (pizza)
        fig4, canvas4 = create_canvas()
        ax4 = fig4.add_subplot(111)
        decadas = [x[0] for x in metrics['decadas']]
        dec_qtd = [x[1] for x in metrics['decadas']]
        ax4.pie(dec_qtd, labels=decadas, autopct='%1.1f%%', startangle=140)
        ax4.set_title("Publicações por Década")
        tab4 = QWidget()
        l4 = QVBoxLayout(tab4)
        l4.addWidget(canvas4)
        subtabs.addTab(tab4, "Décadas")
        self.figures["Décadas"] = fig4

        self.tabs.addTab(widget, "Gráficos")

        # Agora conectamos o botão à função de exportação
        export_plot_button.clicked.connect(self.export_current_plot)

    def export_current_plot(self):
        # Pega a aba de gráficos (assumindo que é a última aba adicionada)
        

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
        save_folder.mkdir(parents=True, exist_ok=True)  # cria a pasta se não existir

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
       #    QMessageBox.critical(self, "Erro", f"Falha ao carregar dados: {str(e)}")
            return

        if not metrics or not metrics['publicacoes_completas']:
            QMessageBox.warning(self, "Aviso", "Nenhum dado encontrado no arquivo BibTeX.")
            return

        self.create_table_tab("Publicações completas", metrics['publicacoes_completas'],
                              ["Ano", "Tipo", "ID", "Autores", "Título", "Publicação"])
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

        if title == "Publicações por ano":
            headers.append("Média por ano em 5 anos")
            table.setColumnCount(len(headers))
            table.setRowCount(len(data))

            valores = [int(row[1]) for row in data]
            anos = [str(row[0]) for row in data]

            for row_idx, row_data in enumerate(data):
                for col_idx, value in enumerate(row_data):
                    item = QTableWidgetItem(str(value))
                    item.setFlags(item.flags() ^ Qt.ItemIsEditable)
                    table.setItem(row_idx, col_idx, item)

                # Calcular média a cada 5 linhas ou no fim
                is_final_row = row_idx == len(data) - 1
                if row_idx % 5 == 4 or is_final_row:
                    start_idx = row_idx - 4 if row_idx >= 4 else 0
                    end_idx = row_idx
                    subset = valores[start_idx:end_idx + 1]
                    intervalo_anos = f"{anos[start_idx]}–{anos[end_idx]}"
                    media_valor = round(sum(subset) / len(subset), 1)
                    texto_media = f"{intervalo_anos} → {media_valor}"

                    media_item = QTableWidgetItem(texto_media)
                    media_item.setFlags(media_item.flags() ^ Qt.ItemIsEditable)
                
                    table.setItem(row_idx, 2, media_item)
                else:
                    table.setItem(row_idx, 2, QTableWidgetItem(""))

        elif title == "Publicações por década":
            headers.append("Média")
            table.setColumnCount(len(headers))
            table.setRowCount(len(data) + 1)  # +1 para linha da média geral

            # Extrair quantidades para cálculo da média
            valores = [int(row[1]) for row in data]

            # Preencher dados normais
            for row_idx, row_data in enumerate(data):
                for col_idx, value in enumerate(row_data):
                    item = QTableWidgetItem(str(value))
                    item.setFlags(item.flags() ^ Qt.ItemIsEditable)
                    table.setItem(row_idx, col_idx, item)

            # Calcular média geral
            media_geral = round(sum(valores) / len(valores), 2) if valores else 0

            # Linha extra para mostrar a média
            media_text_item = QTableWidgetItem("Média geral")
            media_text_item.setFlags(media_text_item.flags() ^ Qt.ItemIsEditable)
            table.setItem(len(data), 0, media_text_item)

            media_valor_item = QTableWidgetItem(str(media_geral))
            media_valor_item.setFlags(media_valor_item.flags() ^ Qt.ItemIsEditable)
            table.setItem(len(data), len(headers)-1, media_valor_item)

            # Preencher as células vazias da média na coluna média com "-"
            for row_idx in range(len(data)):
                table.setItem(row_idx, len(headers)-1, QTableWidgetItem("-"))

        else:
            table.setRowCount(len(data))
            table.setColumnCount(len(headers))
            table.setHorizontalHeaderLabels(headers)

            for row_idx, row_data in enumerate(data):
                for col_idx, value in enumerate(row_data):
                    item = QTableWidgetItem(str(value))
                    item.setFlags(item.flags() ^ Qt.ItemIsEditable)
                    table.setItem(row_idx, col_idx, item)

        table.setHorizontalHeaderLabels(headers)
        table.resizeColumnsToContents()
        self.tabs.addTab(widget, title)

def analyze_bibtex():
    input_file = Path.home() / "Desktop/PYMT/references.bib"

    if not input_file.exists():
        raise FileNotFoundError(f"Arquivo '{input_file}' não encontrado.")

    with open(input_file, 'r', encoding='utf-8') as bibtex_file:
        parser = bibtexparser.bparser.BibTexParser(common_strings=True)
        bib_database = bibtexparser.load(bibtex_file, parser=parser)

    metrics = {
        'publicacoes_completas': [],
        'autores_completos': defaultdict(int),
        'publicacoes_por_ano': defaultdict(int),
        'tipos_publicacao': defaultdict(int),
        'decadas': defaultdict(int),
    }

    anos = []

    for entry in bib_database.entries:
        ano = entry.get('year', 'N/A')
        tipo = entry.get('ENTRYTYPE', 'N/A')
        id_pub = entry.get('ID', 'N/A')
        autores = entry.get('author', 'N/A')
        titulo = entry.get('title', 'N/A')
        publicacao = entry.get('journal', entry.get('booktitle', 'N/A'))

        metrics['publicacoes_completas'].append(
            [ano, tipo, id_pub, autores, titulo, publicacao]
        )

        try:
            ano_int = int(ano)
            metrics['publicacoes_por_ano'][ano_int] += 1
            anos.append(ano_int)
            decade = (ano_int // 10) * 10
            metrics['decadas'][f"{decade}-{decade+9}"] += 1
        except:
            pass

        metrics['tipos_publicacao'][tipo] += 1
        if autores and autores.strip().upper() != "N/A":
            metrics['autores_completos'][autores] += 1

    def dict_to_sorted_list(d, key_format=str):
        return [[key_format(k), v] for k, v in sorted(d.items())]

    # Agrupamento de publicações por intervalo de 5 anos
    publicacoes_por_5anos = defaultdict(int)
    publicacoes_5anos_lista = []

    if anos:
        min_ano = min(anos)
        max_ano = max(anos)
        for ano in anos:
            inicio_intervalo = (ano // 5) * 5
            fim_intervalo = inicio_intervalo + 4
            chave_intervalo = f"{inicio_intervalo}-{fim_intervalo}"
            publicacoes_por_5anos[chave_intervalo] += 1

        # Ordenar os intervalos
        publicacoes_5anos_lista = sorted(
            [[intervalo, qtd] for intervalo, qtd in publicacoes_por_5anos.items()],
            key=lambda x: int(x[0].split("-")[0])
        )

        media_5anos = round(statistics.mean(publicacoes_por_5anos.values()), 2)
    else:
        media_5anos = "N/A"




    result = {
        'publicacoes_completas': metrics['publicacoes_completas'],
        'publicacoes_por_ano': dict_to_sorted_list(metrics['publicacoes_por_ano'], int),
        'autores_completos': sorted(metrics['autores_completos'].items(), key=lambda x: x[1], reverse=True),
        'tipos_publicacao': dict_to_sorted_list(metrics['tipos_publicacao']),
        'decadas': dict_to_sorted_list(metrics['decadas']),
        'estatisticas_gerais': [
            ["Total de publicações", len(metrics['publicacoes_completas'])],
            ["Total de autores distintos", len(metrics['autores_completos'])],
            ["Anos distintos", len(set(anos))],
            ["Ano mais antigo", min(anos) if anos else "N/A"],
            ["Ano mais recente", max(anos) if anos else "N/A"],
       
        ],
        'publicacoes_por_5anos': publicacoes_5anos_lista
    }

    return result


if __name__ == "__main__":
    app = QApplication(sys.argv)
    viewer = BibtexViewer()
    viewer.show()
    sys.exit(app.exec_())
