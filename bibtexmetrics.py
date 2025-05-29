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

class BibtexViewer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Visualizador de BibTeX")
        self.resize(1000, 600)
        self.initUI()
    def initUI(self):
        central_widget = QWidget()
        main_layout = QVBoxLayout()
        central_widget.setLayout(main_layout)
        self.setCentralWidget(central_widget)

        # Botões
        button_layout = QHBoxLayout()

        self.update_button = QPushButton("Atualizar Tabelas")
        self.update_button.clicked.connect(self.reload_data)
       # button_layout.addWidget(self.update_button)

        self.export_button = QPushButton("Exportar Tabela Atual")
        self.export_button.clicked.connect(self.exportar_tabela_atual)
        button_layout.addWidget(self.export_button)

       

        # Abas com tabelas
        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)
        main_layout.addLayout(button_layout)
        self.reload_data()
   
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
            QMessageBox.critical(self, "Erro", f"Falha ao carregar dados: {str(e)}")
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

    def create_table_tab(self, title, data, headers):
        widget = QWidget()
        layout = QVBoxLayout()
        table = QTableWidget()
        layout.addWidget(table)
        widget.setLayout(layout)

        table.setRowCount(len(data))
        table.setColumnCount(len(headers))
        table.setHorizontalHeaderLabels(headers)

        for row_idx, row_data in enumerate(data):
            for col_idx, value in enumerate(row_data):
                item = QTableWidgetItem(str(value))
                item.setFlags(item.flags() ^ Qt.ItemIsEditable)
                table.setItem(row_idx, col_idx, item)

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
            ["Média de publicações por ano", round(statistics.mean(metrics['publicacoes_por_ano'].values()), 2) if anos else "N/A"],
            ["Mediana de publicações por ano", statistics.median(metrics['publicacoes_por_ano'].values()) if anos else "N/A"],
        ]
    }

    return result


if __name__ == "__main__":
    app = QApplication(sys.argv)
    viewer = BibtexViewer()
    viewer.show()
    sys.exit(app.exec_())
