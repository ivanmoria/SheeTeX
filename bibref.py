import sys
import re
import os
import csv
from pathlib import Path
import shared_data 
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QTextEdit, QPushButton,
    QVBoxLayout, QMessageBox, QTableWidget, QTableWidgetItem,
    QTabWidget, QHBoxLayout, QFileDialog,
)
from PyQt5.QtCore import Qt, QStandardPaths
from bibtexmetrics import BibtexViewer
from processador_referencias import extrair_campos_apa, extrair_autores_completos
import importlib
import config

class APA2BibtexWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Bibtex from APA column Ref")
        self.initUI()
        self.salvar_bibtex()
        self.carregar_excel()
    def initUI(self):
        main_layout = QVBoxLayout()

        # Campo oculto para entrada de texto, se precisar (não adicionado visualmente)
        self.input_edit = QTextEdit()
   
        # Tabela APA
        self.tabela_apa = QTableWidget()
        self.tabela_apa.setColumnCount(3)
        self.tabela_apa.setHorizontalHeaderLabels(["Nome", "Ano", "Referência"])
        self.tabela_apa.setColumnWidth(2, 100)

        # Botão Salvar BibTeX
        self.btn_salvar_bibtex = QPushButton("Salvar BibTeX")
        self.btn_salvar_bibtex.clicked.connect(self.salvar_bibtex)

        # Botão para carregar arquivo .bib manualmente
        self.btn_carregar_bib = QPushButton("Carregar arquivo .bib manualmente")
        self.btn_carregar_bib.clicked.connect(self.carregar_bibtex_manual)

        # Layout horizontal para os dois botões
        botoes_bibtex_layout = QHBoxLayout()
        botoes_bibtex_layout.addWidget(self.btn_salvar_bibtex)
        botoes_bibtex_layout.addWidget(self.btn_carregar_bib)


        # Label e área de texto para resultado BibTeX
        main_layout.addWidget(QLabel("Resultado BibTeX:"))
        self.output_edit = QTextEdit()
        self.output_edit.setReadOnly(False)
        main_layout.addWidget(self.output_edit)



        # Adiciona o layout dos botões ao layout principal
        main_layout.addLayout(botoes_bibtex_layout)




        self.setLayout(main_layout)
        self.resize(900, 700)


    def carregar_bibtex_manual(self):
        arquivo_bib, _ = QFileDialog.getOpenFileName(
            self,
            "Selecionar arquivo .bib",
            str(Path.home()),
            "Arquivos BibTeX (*.bib)"
        )

        if arquivo_bib:
            try:
                with open(arquivo_bib, 'r', encoding='utf-8') as f:
                    conteudo_bib = f.read()
                self.output_edit.setPlainText(conteudo_bib)
                QMessageBox.information(self, "BibTeX carregado", "Arquivo .bib carregado com sucesso.")
                self.salvar_bibtex()
            except Exception as e:
                QMessageBox.critical(self, "Erro", f"Erro ao carregar o arquivo .bib:\n{e}")

     

    def carregar_excel(self):
        # Usa sempre self.sheet_csv_url, que é uma URL fixa
        importlib.reload(config)
        url_csv = config.sheet_csv_url

        try:
            df = pd.read_csv(url_csv)
            if "Ref" not in df.columns:
                QMessageBox.critical(self, "Erro", "A coluna 'Ref' não foi encontrada na planilha.")
                return
            referencias = "\n\n".join(
                str(ref).strip() for ref in df["Ref"]
                if pd.notna(ref) and not str(ref).strip().startswith(".")
            )

            self.input_edit.setPlainText(referencias)
            self.converter()
            self.salvar_bibtex()
 
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao ler a planilha do Google Sheets:\n{e}")

    def exportar_tabela(self, tabela):
        desktop = QStandardPaths.writableLocation(QStandardPaths.DesktopLocation)
        pasta_pymt = os.path.join(desktop, "PYMT")
        os.makedirs(pasta_pymt, exist_ok=True)

        nome_base = self.tabs.tabText(self.tabs.currentIndex()) or "tabela_exportada"
        nome_base = nome_base.replace(" ", "_")
        extensao = ".csv"
        caminho_arquivo = os.path.join(pasta_pymt, nome_base + extensao)

        contador = 1
        while os.path.exists(caminho_arquivo):
            caminho_arquivo = os.path.join(pasta_pymt, f"{nome_base}_{contador}{extensao}")
            contador += 1

        try:
            with open(caminho_arquivo, 'w', encoding='utf-8', newline='') as f:
                writer = csv.writer(f)
                colunas = tabela.columnCount()
                linhas = tabela.rowCount()

                headers = [
                    tabela.horizontalHeaderItem(c).text() if tabela.horizontalHeaderItem(c) else f"Coluna {c+1}"
                    for c in range(colunas)
                ]
                writer.writerow(headers)

                for r in range(linhas):
                    linha = [
                        tabela.item(r, c).text() if tabela.item(r, c) else ''
                        for c in range(colunas)
                    ]
                    writer.writerow(linha)

            QMessageBox.information(self, "Exportar tabela", f"Tabela exportada com sucesso para:\n{caminho_arquivo}")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Falha ao exportar tabela:\n{str(e)}")

    def converter(self):
        entrada = self.input_edit.toPlainText().strip()
        if not entrada:
            QMessageBox.warning(self, "Aviso", "Não há referências para converter.")
            return
        entradas = re.split(r'\n{2,}', entrada)
        entradas = [e for e in entradas if not e.strip().startswith(".")]

        referencias_separadas = []
        for entrada in entradas:
            partes = [r.strip() for r in entrada.split('\n\n') if r.strip()]
            referencias_separadas.extend(partes)

        self.tabela_apa.setRowCount(len(referencias_separadas))
        for i, ref in enumerate(referencias_separadas):
            # Regex que captura nome, ano e o resto da referência após o ano
            match = re.search(r'^(.*?)\s*\((\d{4})\)\.?\s*(.*)$', ref)
            if match:
                nome = match.group(1).strip()
                ano = match.group(2)
                titulo = match.group(3).strip()  # Título após o ano e possível ponto removido
            else:
                nome = ""
                ano = ""
                titulo = ""

            # Você pode juntar nome e título se quiser exibir separados ou em uma coluna só
            # Por exemplo, na coluna 'Nome' deixar só o nome, e na 'Referência' o título completo
            item_nome = QTableWidgetItem(nome)
            item_ano = QTableWidgetItem(ano)
            item_ref = QTableWidgetItem(titulo if titulo else ref)  # se título vazio, usa ref inteiro

            item_ref.setTextAlignment(Qt.AlignTop)
            item_ref.setToolTip(ref)

            self.tabela_apa.setItem(i, 0, item_nome)
            self.tabela_apa.setItem(i, 1, item_ano)
            self.tabela_apa.setItem(i, 2, item_ref)

        self.tabela_apa.resizeColumnsToContents()
        self.tabela_apa.resizeRowsToContents()

        bibtex_total = ""
        lista_campos = []
        for idx, entrada_individual in enumerate(referencias_separadas, 1):
            campos = extrair_campos_apa(entrada_individual)
            lista_campos.append(campos)

            bibtex = f"@article{{ref{idx},\n"
            for campo, valor in campos.items():
                if valor:
                    bibtex += f"  {campo} = {{{valor}}},\n"
            bibtex += f"  note = {{{entrada_individual.strip()}}}\n}}\n\n"
            bibtex_total += bibtex

        self.df_bibtex = pd.DataFrame(lista_campos)

        self.output_edit.setPlainText(bibtex_total.strip())
        # self.tabs.setCurrentWidget(self.tab_bibtex)

        shared_data.df_bibtex = self.output_edit


    def exportar_csv(self):
        desktop = QStandardPaths.writableLocation(QStandardPaths.DesktopLocation)
        pasta_pymt = os.path.join(desktop, "PYMT")
        os.makedirs(pasta_pymt, exist_ok=True)
        arquivo_csv = os.path.join(pasta_pymt, "APA_referencias.csv")

        try:
            with open(arquivo_csv, 'w', encoding='utf-8', newline='') as f:
                writer = csv.writer(f)
                writer.writerow(["Nome", "Ano", "Referência"])
                linhas = self.tabela_apa.rowCount()
                colunas = self.tabela_apa.columnCount()

                for r in range(linhas):
                    linha = [
                        self.tabela_apa.item(r, c).text() if self.tabela_apa.item(r, c) else ''
                        for c in range(colunas)
                    ]
                    writer.writerow(linha)

            QMessageBox.information(self, "Exportar CSV", f"Arquivo CSV exportado:\n{arquivo_csv}")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao exportar CSV:\n{str(e)}")

    def salvar_bibtex(self):
        # Define a pasta Desktop/PYMT dentro do diretório home do usuário
        pasta = Path.home() / "Desktop" / "PYMT"
        pasta.mkdir(parents=True, exist_ok=True)  # cria a pasta se não existir

        arquivo = pasta / "references.bib"  # nome do arquivo a salvar

        texto = self.output_edit.toPlainText()
        if not texto.strip():
            return

        try:
            with open(arquivo, "w", encoding="utf-8") as f:
                f.write(texto)
            # Opcional: mostrar confirmação
            # QMessageBox.information(self, "BibTeX salvo", f"Arquivo salvo em:\n{arquivo}")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao salvar BibTeX:\n{str(e)}")

def main():
    app = QApplication(sys.argv)
    widget = APA2BibtexWidget()
    widget.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
