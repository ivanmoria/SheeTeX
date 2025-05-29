import sys
import re
import os
import csv
from pathlib import Path
import shared_data 
import pandas as pd
from collections import Counter
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QTextEdit, QPushButton,
    QVBoxLayout, QMessageBox, QDialog, QListWidget, QListWidgetItem,
    QSplitter, QTableWidget, QTableWidgetItem,
    QTabWidget, QHBoxLayout, QFileDialog
)
from PyQt5.QtCore import Qt, QStandardPaths
from bibtexmetrics import BibtexViewer

def extrair_campos_apa(texto):
    campos = {
        'author': '', 'year': '', 'title': '',
        'journal': '', 'booktitle': '', 'number': '',
        'pages': '', 'doi': ''
    }
    texto = ' '.join(texto.split())
    match_ano = re.search(r'\((\d{4})\)', texto)
    if not match_ano:
        return campos
    campos['year'] = match_ano.group(1)
    inicio_ano = match_ano.start()
    fim_ano = match_ano.end()
    campos['author'] = texto[:inicio_ano].strip().rstrip('.')
    depois = texto[fim_ano:].strip()
    match_titulo = re.match(r'(.+?)\.\s+(.*)', depois)
    if not match_titulo:
        return campos
    campos['title'] = match_titulo.group(1).strip()
    resto = match_titulo.group(2)
    if ' In ' in resto or resto.startswith('In '):
        match_in = re.search(r'\bIn\b (.+?)(?:\.|$)', resto)
        if match_in:
            campos['booktitle'] = match_in.group(1).strip()
        match_paginas = re.search(r'\(pp\.\s*(\d+-\d+)\)', texto)
        if match_paginas:
            campos['pages'] = match_paginas.group(1)
    else:
        match_journal = re.match(r'([^.]+)\. (.+)', resto)
        if match_journal:
            campos['journal'] = match_journal.group(1).strip()
            resto = match_journal.group(2)
    match_number = re.search(r'(\d+\s*\(\d+\))', resto)
    if match_number:
        campos['number'] = match_number.group(1)
    match_pages = re.search(r'(\d{1,4}-\d{1,4})', resto)
    if match_pages:
        campos['pages'] = match_pages.group(1)
    match_doi = re.search(r'(10\.\d{4,9}/[-._;()/:A-Z0-9]+)', texto, re.IGNORECASE)
    if match_doi:
        campos['doi'] = match_doi.group(1)
    return campos

def extrair_autores_completos(campo_autor):
    possiveis_autores = re.split(r';|\.\s+', campo_autor)
    autores_extraidos = []
    for autor in possiveis_autores:
        autor = autor.strip()
        if not autor:
            continue
        match = re.match(r'([A-Z][a-zA-ZÀ-ÿ\-]+),\s*([A-Z][a-zA-ZÀ-ÿ\-. ]+)', autor)
        if match:
            nome = match.group(0).strip().rstrip('.')
            autores_extraidos.append(nome)
        else:
            autores_extraidos.append(autor.strip().rstrip('.'))
    return autores_extraidos

class MetricsDialog(QDialog):
    def __init__(self, autores_contagem, anos_contagem):
        super().__init__()
        self.setWindowTitle("Métricas de Referências")

        lista_autores = QListWidget()
        for autor, qtd in autores_contagem.most_common(10):
            item = QListWidgetItem(f"{autor} — {qtd} ocorrência(s)")
            lista_autores.addItem(item)
       

class APA2BibtexWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Conversor APA → BibTeX (de planilha)")
        self.initUI()
        self.carregar_excel()  # Carrega automaticamente do Google Sheets
        self.converter()       # Converte automaticamente
        self.salvar_bibtex()

   
    def initUI(self):
            main_layout = QVBoxLayout()
            self.input_edit = QTextEdit()

            self.tabs = QTabWidget()

            # Aba APA
            self.tab_apa = QWidget()
            tab_apa_layout = QVBoxLayout()
            self.tabela_apa = QTableWidget()
            self.tabela_apa.setColumnCount(3)
            self.tabela_apa.setHorizontalHeaderLabels(["Nome", "Ano", "Referência"])
            self.tabela_apa.setColumnWidth(2, 600)  # largura para exibir referência completa
            tab_apa_layout.addWidget(self.tabela_apa)
            self.btn_exportar_csv = QPushButton("Exportar CSV")
            self.btn_exportar_csv.clicked.connect(self.exportar_csv)
            tab_apa_layout.addWidget(self.btn_exportar_csv)
            self.tab_apa.setLayout(tab_apa_layout)
            self.tabs.addTab(self.tab_apa, "Visualização APA em Tabela")

            # Aba BibTeX
            self.tab_bibtex = QWidget()
            tab_bibtex_layout = QVBoxLayout()
            tab_bibtex_layout.addWidget(QLabel("Resultado BibTeX:"))
            self.output_edit = QTextEdit()
            self.output_edit.setReadOnly(True)
            tab_bibtex_layout.addWidget(self.output_edit)
            self.btn_salvar_bibtex = QPushButton("Salvar BibTeX")
            self.btn_salvar_bibtex.clicked.connect(self.salvar_bibtex)
            tab_bibtex_layout.addWidget(self.btn_salvar_bibtex)
            self.tab_bibtex.setLayout(tab_bibtex_layout)
            self.tabs.addTab(self.tab_bibtex, "BibTeX")

            # --- Nova aba para o BibtexViewer ---
            self.tab_bibtexviewer = QWidget()
            tab_bibtexviewer_layout = QVBoxLayout()

            # Instancia o BibtexViewer e adiciona na aba
            self.bibtex_viewer = BibtexViewer()  
            tab_bibtexviewer_layout.addWidget(self.bibtex_viewer)

            self.tab_bibtexviewer.setLayout(tab_bibtexviewer_layout)
            self.tabs.addTab(self.tab_bibtexviewer, "Bibtex Metrics Viewer")

            main_layout.addWidget(self.tabs)


            self.setLayout(main_layout)
            self.resize(900, 700)


    def exportar_tabela(self, tabela):
        # Define o caminho da pasta "PYMT" na área de trabalho
        desktop = QStandardPaths.writableLocation(QStandardPaths.DesktopLocation)
        pasta_pymt = os.path.join(desktop, "PYMT")
        os.makedirs(pasta_pymt, exist_ok=True)

        # Nome base do arquivo conforme aba ativa
        nome_base = self.tabs.tabText(self.tabs.currentIndex()) or "tabela_exportada"
        nome_base = nome_base.replace(" ", "_")  # Substitui espaços por underscores
        extensao = ".csv"
        caminho_arquivo = os.path.join(pasta_pymt, nome_base + extensao)

        # Garante que o nome do arquivo seja único
        contador = 1
        while os.path.exists(caminho_arquivo):
            caminho_arquivo = os.path.join(pasta_pymt, f"{nome_base}_{contador}{extensao}")
            contador += 1

        # Tenta escrever os dados da tabela no arquivo CSV
        try:
            with open(caminho_arquivo, 'w', encoding='utf-8', newline='') as f:
                writer = csv.writer(f)
                colunas = tabela.columnCount()
                linhas = tabela.rowCount()

                # Escreve os cabeçalhos
                headers = [
                    tabela.horizontalHeaderItem(c).text() if tabela.horizontalHeaderItem(c) else f"Coluna {c+1}"
                    for c in range(colunas)
                ]
                writer.writerow(headers)

                # Escreve os dados da tabela
                for r in range(linhas):
                    linha = [
                        tabela.item(r, c).text() if tabela.item(r, c) else ''
                        for c in range(colunas)
                    ]
                    writer.writerow(linha)

            # Mensagem de sucesso
            QMessageBox.information(self, "Exportar tabela", f"Tabela exportada com sucesso para:\n{caminho_arquivo}")
        except Exception as e:
            # Mensagem de erro
            QMessageBox.critical(self, "Erro", f"Falha ao exportar tabela:\n{str(e)}")


    def carregar_excel(self):
        try:
            url = "https://docs.google.com/spreadsheets/d/1XuGWm_gDG5edw9YkznTQGABTBah1Ptz9lfstoFdGVbA/export?format=csv&gid=49303292"
            df = pd.read_csv(url)
            
            if "Ref" not in df.columns:
                QMessageBox.critical(self, "Erro", "A coluna 'Ref' não foi encontrada na planilha.")
                return
            referencias = "\n\n".join(str(ref).strip() for ref in df["Ref"] if pd.notna(ref))
            self.input_edit.setPlainText(referencias)
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao ler a planilha do Google Sheets:\n{e}")

    def converter(self):
        entrada = self.input_edit.toPlainText().strip()
        if not entrada:
            QMessageBox.warning(self, "Aviso", "Não há referências para converter.")
            return
        entradas = re.split(r'\n{2,}', entrada)

        # Separar referências por enter duplo e criar lista única de referências individuais
        referencias_separadas = []
        for entrada in entradas:
            partes = [r.strip() for r in entrada.split('\n\n') if r.strip()]
            referencias_separadas.extend(partes)

        self.tabela_apa.setRowCount(len(referencias_separadas))
        for i, ref in enumerate(referencias_separadas):
            match = re.search(r'^(.*?)\s*\((\d{4})\)', ref)
            nome = match.group(1).strip() if match else ""
            ano = match.group(2) if match else ""
            item_nome = QTableWidgetItem(nome)
            item_ano = QTableWidgetItem(ano)
            item_ref = QTableWidgetItem(ref)
            item_ref.setTextAlignment(Qt.AlignTop)
            item_ref.setToolTip(ref)

            self.tabela_apa.setItem(i, 0, item_nome)
            self.tabela_apa.setItem(i, 1, item_ano)
            self.tabela_apa.setItem(i, 2, item_ref)

        self.tabela_apa.resizeColumnsToContents()
        self.tabela_apa.resizeRowsToContents()

        bibtex_total = ""
        lista_campos = []  # lista para armazenar dicionários com os campos de cada referência
        for idx, entrada_individual in enumerate(referencias_separadas, 1):
            campos = extrair_campos_apa(entrada_individual)
            lista_campos.append(campos)  # salva o dicionário na lista

            bibtex = f"@article{{ref{idx},\n"
            for campo, valor in campos.items():
                if valor:
                    bibtex += f"  {campo} = {{{valor}}},\n"
            bibtex += f"  note = {{{entrada_individual.strip()}}}\n}}\n\n"
            bibtex_total += bibtex

        # Criar o dataframe final do bibtex antes de salvar
        self.df_bibtex = pd.DataFrame(lista_campos)

        self.output_edit.setPlainText(bibtex_total.strip())
        self.tabs.setCurrentWidget(self.tab_bibtex)

        shared_data.df_bibtex = self.df_bibtex

    def exportar_csv(self):
        desktop = QStandardPaths.writableLocation(QStandardPaths.DesktopLocation)
        pasta_pymt = os.path.join(desktop, "PYMT")
        os.makedirs(pasta_pymt, exist_ok=True)

        try:
            # Itera por todas as abas do QTabWidget
            for i in range(self.tabs.count()):
                nome_aba = self.tabs.tabText(i).replace(" ", "_").replace("/", "_")  # nome da aba seguro para arquivo
                tab_widget = self.tabs.widget(i)

                tabela = tab_widget.findChild(QTableWidget)
                if tabela is None:
                    continue  # pula se não encontrar tabela

                file_path = os.path.join(pasta_pymt, f"{nome_aba}.csv")

                with open(file_path, mode='w', newline='', encoding='utf-8') as arquivo_csv:
                    escritor = csv.writer(arquivo_csv)

                    # Cabeçalhos
                    headers = [
                        tabela.horizontalHeaderItem(col).text() if tabela.horizontalHeaderItem(col) else f"Coluna {col+1}" 
                        for col in range(tabela.columnCount())
                    ]
                    escritor.writerow(headers)

                    # Linhas
                    for row in range(tabela.rowCount()):
                        linha = []
                        for col in range(tabela.columnCount()):
                            item = tabela.item(row, col)
                            linha.append(item.text() if item else "")
                        escritor.writerow(linha)

            QMessageBox.information(self, "Sucesso", f"Arquivos CSV exportados com sucesso em:\n{pasta_pymt}")

        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao salvar os arquivos CSV:\n{e}")
    def salvar_bibtex(self):
        conteudo_bibtex = self.output_edit.toPlainText().strip()

        if not conteudo_bibtex:
            QMessageBox.warning(self, "Aviso", "Não há conteúdo BibTeX para salvar.")
            return

        # Caminho para a Área de Trabalho
        desktop = QStandardPaths.writableLocation(QStandardPaths.DesktopLocation)

        # Caminho para a pasta PYMT dentro da área de trabalho
        pasta_pymt = os.path.join(desktop, "PYMT")

        # Cria a pasta se ela não existir
        os.makedirs(pasta_pymt, exist_ok=True)

        # Caminho completo do arquivo BibTeX a ser salvo
        nome_arquivo = os.path.join(pasta_pymt, "references.bib")

        try:
            with open(nome_arquivo, 'w', encoding='utf-8') as f:
                f.write(conteudo_bibtex)

        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao salvar arquivo:\n{e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    janela = APA2BibtexWidget()
    janela.show()
    sys.exit(app.exec_())
