# interface.py
import os
import re
import pandas as pd
from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QTableWidget, QLabel,
    QPushButton, QMessageBox, QLineEdit, QTextEdit, QSplitter, QTabWidget,
    QMenuBar, QAction
)
from PyQt5.QtCore import Qt, QStandardPaths
from PyQt5.QtGui import QIcon
from bibref import APA2BibtexWidget  # Supondo que você tenha esse módulo

class GoogleSheetsViewer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowIcon(QIcon("we.icns")) 
        self.setWindowTitle("Visualizador de Planilha Google")
        self.setGeometry(100, 100, 1100, 700)

        self.sheet_csv_url = (
            "https://docs.google.com/spreadsheets/d/1XuGWm_gDG5edw9YkznTQGABTBah1Ptz9lfstoFdGVbA/export?format=csv&gid=49303292"
        )
        self.dataframe = pd.DataFrame()
        
        self.init_ui()
        self.load_data()
        self.refresh_data()

    def init_ui(self):
        # Central widget e layout principal
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        main_layout = QVBoxLayout(self.central_widget)
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(10)

        # Linha para URL da planilha CSV
        url_layout = QHBoxLayout()
        url_label_text = QLabel("URL da Planilha .csv utilizada neste ambiente:")
        self.url_lineedit = QLineEdit()
        self.url_lineedit.setText(self.sheet_csv_url)
        self.url_lineedit.setToolTip("Edite este campo para usar outra planilha CSV pública")
        self.refresh_button = QPushButton("🔄 Atualizar URL e Dados")
        self.refresh_button.clicked.connect(self.refresh_data)
        url_layout.addWidget(url_label_text)
        url_layout.addWidget(self.url_lineedit)
        url_layout.addWidget(self.refresh_button)
        main_layout.addLayout(url_layout)

        # Abas principais
        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)

        # Aba: Visualizador de Planilha
        self.visualizador_widget = QWidget()
        visualizador_layout = QHBoxLayout(self.visualizador_widget)
        visualizador_layout.setContentsMargins(0, 0, 0, 0)
        visualizador_layout.setSpacing(5)

        # Painel esquerdo do visualizador
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(5)

        self.region_label = QLabel("Visualizando: Geral (Todas)")
        left_layout.addWidget(self.region_label)

        self.metrics_label = QLabel("")
        self.metrics_label.setStyleSheet("font-style: italic; color: gray;")
        left_layout.addWidget(self.metrics_label)

        self.table = QTableWidget()
        self.table.verticalHeader().setVisible(False)
        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(False)
        self.table.setSelectionBehavior(QTableWidget.SelectItems)
        self.table.horizontalHeader().setStretchLastSection(True)
        left_layout.addWidget(self.table)

        # Botões de ação
        button_layout = QHBoxLayout()
        button_layout.setSpacing(10)

        self.refresh_button = QPushButton("🔄")
        self.refresh_button.clicked.connect(self.refresh_data)
        button_layout.addWidget(self.refresh_button)

        self.reload_button = QPushButton("🔄 Recarregar Dados")
        self.reload_button.clicked.connect(self.load_data)
        button_layout.addWidget(self.reload_button)

        self.export_button = QPushButton("📀 Exportar CSV")
        self.export_button.clicked.connect(self.export_to_csv)
        button_layout.addWidget(self.export_button)

        self.export_authors_button = QPushButton("📥 Exportar Autores + Afiliações")
        self.export_authors_button.clicked.connect(self.export_authors_affiliations)
        button_layout.addWidget(self.export_authors_button)

        self.export_refs_button = QPushButton("📜 Exportar Referências Completas")
        self.export_refs_button.clicked.connect(self.export_full_references)
        button_layout.addWidget(self.export_refs_button)

        left_layout.addLayout(button_layout)

        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(left_widget)
        visualizador_layout.addWidget(splitter)

        self.tabs.addTab(self.visualizador_widget, "Visualizador de Planilha")

        # Aba: Métricas
        self.aba_metricas = QWidget()
        layout_metricas = QVBoxLayout(self.aba_metricas)
        layout_metricas.setContentsMargins(10, 10, 10, 10)
        layout_metricas.setSpacing(10)

        self.metricas_textedit = QTextEdit()
        self.metricas_textedit.setReadOnly(True)
        layout_metricas.addWidget(self.metricas_textedit)

        self.export_metricas_button = QPushButton("💾 Exportar Métricas para CSV")
        self.export_metricas_button.clicked.connect(self.exportar_metricas_para_csv)
        layout_metricas.addWidget(self.export_metricas_button)

        self.tabs.addTab(self.aba_metricas, " Métricas")

        # Aba: Bibtex Metrics Viewer
        self.tab_bibtexviewer = QWidget()
        tab_bibtexviewer_layout = QVBoxLayout()
        self.bibtex_viewer = APA2BibtexWidget()
        tab_bibtexviewer_layout.addWidget(self.bibtex_viewer)
        self.tab_bibtexviewer.setLayout(tab_bibtexviewer_layout)
        self.tabs.addTab(self.tab_bibtexviewer, "Bibtex Metrics Viewer")

    def load_data(self):
        # Exemplo básico: carregar CSV da URL
        try:
            self.dataframe = pd.read_csv(self.url_lineedit.text())
            self.populate_table()
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao carregar CSV:\n{e}")

    def refresh_data(self):
        # Pode ser usado para atualizar filtros, labels etc.
        self.region_label.setText("Visualizando: Geral (Todas)")

    def populate_table(self):
        df = self.dataframe
        if df.empty:
            return
        self.table.clear()
        self.table.setColumnCount(len(df.columns))
        self.table.setRowCount(len(df))
        self.table.setHorizontalHeaderLabels(df.columns)

        for i, row in df.iterrows():
            for j, col in enumerate(df.columns):
                item = QTableWidgetItem(str(row[col]))
                self.table.setItem(i, j, item)

    def export_to_csv(self):
        if self.dataframe.empty:
            QMessageBox.warning(self, "Aviso", "Nenhum dado para exportar.")
            return
        desktop = QStandardPaths.writableLocation(QStandardPaths.DesktopLocation)
        path = os.path.join(desktop, "exported_data.csv")
        try:
            self.dataframe.to_csv(path, index=False)
            QMessageBox.information(self, "Sucesso", f"Arquivo salvo em:\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao salvar arquivo:\n{e}")

    def export_authors_affiliations(self):
        # (Implementação conforme seu código original)
        pass

    def export_full_references(self):
        # (Implementação conforme seu código original)
        pass

    def exportar_metricas_para_csv(self):
        # (Implementação conforme seu código original)
        pass
