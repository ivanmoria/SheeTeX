import sys
import re
import time
import os
import os
import csv
from PyQt5.QtWidgets import QPushButton, QMessageBox
from PyQt5.QtCore import QStandardPaths

from config import sheet_csv_url
import traceback
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout,
    QTableWidget, QTableWidgetItem,QSplashScreen, QPushButton, QMessageBox,
    QFileDialog, QHBoxLayout, QMenuBar, QLabel, QAction,QLineEdit,
    QTextEdit, QSplitter, QTabWidget)
from PyQt5.QtCore import Qt, QTimer, QStandardPaths
from PyQt5.QtGui import QColor, QFont, QPixmap, QIcon
from processador_referencias import extrair_campos_apa, extrair_autores_completos
from bibref import APA2BibtexWidget 
from metricas import calcular_metricas,  exportar_metricas_texto_para_csv, calcular_metricas_detalhadas
from bibtexmetrics import BibtexViewer
from bib import     processar_entrada_bibtex, processar_visualizacao_formatada, processar_estatisticas_bibtex
class GoogleSheetsViewer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowIcon(QIcon("we.icns")) 
        self.setWindowTitle("Visualizador de Planilha Google")
        self.setGeometry(100, 100, 1100, 700)
        self.sheet_csv_url = (
"https://docs.google.com/spreadsheets/d/1Kf5_DYjTErSNpVjPoKG0HIyKfAhBfuVM_Qt4a0EOMYQ/export?format=csv&gid=49303292")
        
        

        self.dataframe = pd.DataFrame()
        self.selected_region = None
        self.init_ui()
        self.load_data()
        self.refresh_data()

        self.visualizacao_bibtex_widget.carregar_excel()


        self.estatisticas_bibtex_widget.reload_data()
            
    def init_ui(self):
        # --- Widget central e layout principal ---
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        main_layout = QVBoxLayout(self.central_widget)
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(10)

        # --- Linha para URL da planilha CSV ---
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

        # --- Abas principais ---
        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)

        # === Aba: Visualizador de Planilha ===
        self.visualizador_widget = QWidget()
        visualizador_layout = QHBoxLayout(self.visualizador_widget)
        visualizador_layout.setContentsMargins(0, 0, 0, 0)
        visualizador_layout.setSpacing(5)

        # --- Menu de filtros ---
        menubar = QMenuBar(self)
        self.setMenuBar(menubar)
        self.region_menu = menubar.addMenu("Filtrar por Região")
        self.author_menu = menubar.addMenu("Filtrar por Autor")

        # --- Painel esquerdo do visualizador ---
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(5)

        self.region_label = QLabel("Visualizando: Geral (Todas)")
        self.region_label.setContentsMargins(0, 0, 0, 0)
        left_layout.addWidget(self.region_label)

        self.metrics_label = QLabel("hh")
        self.metrics_label.setStyleSheet("font-style: italic; color: gray;")
        self.metrics_label.setContentsMargins(0, 0, 0, 0)
        left_layout.addWidget(self.metrics_label)

        # --- Tabela de dados ---
        self.table = QTableWidget()
        self.table.verticalHeader().setVisible(False)
        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(False)
        self.table.setSelectionBehavior(QTableWidget.SelectItems)
        self.table.horizontalHeader().setStretchLastSection(True)
        left_layout.addWidget(self.table)

        # --- Botões de ação ---
        button_layout = QHBoxLayout()
        button_layout.setSpacing(10)

        self.refresh_button = QPushButton("🔄")
        self.refresh_button.clicked.connect(self.refresh_data)
      #  button_layout.addWidget(self.refresh_button)

        self.reload_button = QPushButton("🔄 Recarregar Dados")
        self.reload_button.clicked.connect(self.load_data)
       # button_layout.addWidget(self.reload_button)

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

        # --- Divisor para visualização lateral ---
        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(left_widget)
        visualizador_layout.addWidget(splitter)

        self.tabs.addTab(self.visualizador_widget, "Visualizador de Planilha")

        # === Aba: Métricas ===
        self.aba_metricas = QWidget()
        layout_metricas = QVBoxLayout(self.aba_metricas)
        layout_metricas.setContentsMargins(10, 10, 10, 10)
        layout_metricas.setSpacing(10)

        # Criando o QTabWidget interno para subabas dentro da aba Métricas
        self.subtabs_metricas = QTabWidget()
        layout_metricas.addWidget(self.subtabs_metricas)

        # Subaba 1: Visão Geral
        self.subaba_visao_geral = QWidget()
        layout_visao_geral = QVBoxLayout(self.subaba_visao_geral)
        layout_visao_geral.setContentsMargins(10, 10, 10, 10)
        layout_visao_geral.setSpacing(10)

        self.metricas_textedit = QTextEdit()
        self.metricas_textedit.setReadOnly(True)
        layout_visao_geral.addWidget(self.metricas_textedit)

        self.export_metricas_button = QPushButton("💾 Exportar Métricas para CSV")
        self.export_metricas_button.clicked.connect(self.exportar_metricas_para_csv)
        layout_visao_geral.addWidget(self.export_metricas_button)

        

        # Subaba 2: Detalhes
        self.subaba_detalhes = QWidget()
        layout_detalhes = QVBoxLayout(self.subaba_detalhes)
        layout_detalhes.setContentsMargins(10, 10, 10, 10)
        layout_detalhes.setSpacing(10)

        self.metricas_detalhes_textedit = QTextEdit()
        self.metricas_detalhes_textedit.setReadOnly(True)
        layout_detalhes.addWidget(self.metricas_detalhes_textedit)

        # Botão para salvar o conteúdo da aba detalhes em CSV
        self.btn_salvar_csv = QPushButton("Salvar CSV")
        layout_detalhes.addWidget(self.btn_salvar_csv)

        self.btn_salvar_csv.clicked.connect(self.salvar_detalhes_em_csv)


        self.subtabs_metricas.addTab(self.subaba_detalhes, " Detalhes")
        self.subtabs_metricas.addTab(self.subaba_visao_geral, " Visão Geral")
        # Adiciona a aba Métricas (com subabas internas) no widget principal
        self.tabs.addTab(self.aba_metricas, " Métricas")
    

        # === Aba: Bibtex com abas internas ===
        self.aba_bibtex = QWidget()
        layout_bibtex = QVBoxLayout(self.aba_bibtex)
        layout_bibtex.setContentsMargins(10, 10, 10, 10)
        layout_bibtex.setSpacing(10)

        # Criando o QTabWidget interno
        self.tabs_bibtex_internos = QTabWidget()
        layout_bibtex.addWidget(self.tabs_bibtex_internos)

        # --- Aba 1: Entrada Bibtex ---
        self.entrada_bibtex_widget = QWidget()
        layout_entrada = QVBoxLayout(self.entrada_bibtex_widget)
        self.entrada_bibtex_textedit = QTextEdit()
        layout_entrada.addWidget(self.entrada_bibtex_textedit)

      
      #  self.tabs_bibtex_internos.addTab(self.entrada_bibtex_widget, "Entrada Bibtex")

        # --- Aba 2: Visualização Formatada ---
        self.visualizacao_bibtex_widget = APA2BibtexWidget ()
        layout_visu = QVBoxLayout(self.visualizacao_bibtex_widget)
        self.visualizacao_bibtex_textedit = QTextEdit()
        self.visualizacao_bibtex_textedit.setReadOnly(True)
        layout_visu.addWidget(self.visualizacao_bibtex_textedit)


      
        # --- Aba 3: Estatísticas Bibtex ---
        self.estatisticas_bibtex_widget = BibtexViewer()
        layout_estat = QVBoxLayout(self.estatisticas_bibtex_widget)
        self.estatisticas_bibtex_textedit = QTextEdit()
        self.estatisticas_bibtex_textedit.setReadOnly(True)
        layout_estat.addWidget(self.estatisticas_bibtex_textedit)

        self.tabs.addTab(self.aba_bibtex, " Bibtex")
        
        self.tabs_bibtex_internos.addTab(self.visualizacao_bibtex_widget, "Bib")

        self.tabs_bibtex_internos.addTab(self.estatisticas_bibtex_widget, "Visualização Formatada")
        # --- Carrega os dados inicialmente ---
        self.refresh_data()
    
    def export_authors_affiliations(self):
        if self.dataframe.empty:
            QMessageBox.warning(self, "Aviso", "Nenhum dado disponível para exportar.")
            return
        if "Autores" not in self.dataframe.columns or "Afiliation" not in self.dataframe.columns:
            QMessageBox.warning(self, "Aviso", "As colunas 'Autores' e 'Afiliation' são necessárias.")
            return
        pairs = []
        for _, row in self.dataframe.iterrows():
            autores = [a.strip() for a in str(row["Autores"]).split(",") if a.strip()]
            raw_afil = str(row["Afiliation"]).replace("\n", " ").strip()
            afiliacoes = re.split(r'\.\s+|\.$', raw_afil)
            afiliacoes = [a.strip() for a in afiliacoes if a.strip()]
            if len(autores) != len(afiliacoes):
                print(f"[⚠️ Aviso] {len(autores)} autores e {len(afiliacoes)} afiliações não coincidem.")
                print("-> Autores:", autores)
                print("-> Afiliacoes:", afiliacoes)
            max_len = max(len(autores), len(afiliacoes))
            for i in range(max_len):
                autor = autores[i] if i < len(autores) else ""
                afil = afiliacoes[i] if i < len(afiliacoes) else ""
                if autor:
                    pairs.append({"Autor": autor, "Afiliation": afil})
        if not pairs:
            QMessageBox.information(self, "Exportação", "Nenhum par Autor/Afiliation encontrado.")
            return
        df_export = pd.DataFrame(pairs)
        desktop = QStandardPaths.writableLocation(QStandardPaths.DesktopLocation)

        # Caminho para a pasta PYMT dentro da área de trabalho
        pasta_pymt = os.path.join(desktop, "PYMT")

        # Cria a pasta se ela não existir
        os.makedirs(pasta_pymt, exist_ok=True)

        # Caminho completo do arquivo a ser salvo
        file_path = os.path.join(pasta_pymt, "autores_afiliacoes.csv")



        try:
            df_export.to_csv(file_path, index=False)
            QMessageBox.information(self, "Sucesso", f"Arquivo salvo com sucesso em:\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao salvar arquivo:\n{e}")
    


    def export_full_references(self):
        try:
            df_raw = pd.read_csv(self.sheet_csv_url)
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao carregar dados crus:\n{e}")
            return
        if "Ref" not in df_raw.columns:
            QMessageBox.warning(self, "Aviso", "Coluna 'Ref' não encontrada.")
            return
        referencias_separadas = []
        for ref_cell in df_raw["Ref"].dropna():
            partes = str(ref_cell).split("\n\n")
            for parte in partes:
                parte_limpa = parte.strip()
                if parte_limpa:
                    referencias_separadas.append(parte_limpa)
        df_export = pd.DataFrame({"Ref": referencias_separadas})


        def extrair_ano(ref_text):
            match = re.search(r"\((\d{4})\)", ref_text)
            if match:
                return match.group(1)
            else:
                return ""

        df_export["Ano"] = df_export["Ref"].apply(extrair_ano)

        desktop = QStandardPaths.writableLocation(QStandardPaths.DesktopLocation)

        # Caminho para a pasta PYMT dentro da área de trabalho
        pasta_pymt = os.path.join(desktop, "PYMT")

        # Cria a pasta se ela não existir
        os.makedirs(pasta_pymt, exist_ok=True)

        # Caminho completo do arquivo a ser salvo
        file_path = os.path.join(pasta_pymt, "referencias_completas.csv")


        try:
            df_export.to_csv(file_path, index=False)
            QMessageBox.information(self, "Sucesso", f"Arquivo salvo com sucesso em:\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao salvar arquivo:\n{e}")
            
    def refresh_data(self):
        try:
            nova_url = self.url_lineedit.text().strip()
            if not nova_url:
                QMessageBox.warning(self, "URL inválida", "Por favor, insira uma URL válida.")
                return

            # Converte para link CSV, atualiza atributo
            self.sheet_csv_url = self.converter_para_link_csv(nova_url)
  

            # 🟡 Atualiza config.py com nova URL
            try:
                with open("config.py", "w", encoding="utf-8") as config_file:
                    config_file.write(f'sheet_csv_url = "{self.sheet_csv_url}"\n')
            except Exception as e:
                QMessageBox.critical(self, "Erro ao salvar URL", f"Não foi possível atualizar config.py:\n{e}")

            # Mostra cursor de carregamento e atualiza UI
            self.setCursor(Qt.WaitCursor)
            QApplication.processEvents()

            # Reseta filtros e seleções
            self.selected_region = None
            if hasattr(self, 'autor_actions'):
                for action in self.autor_actions:
                    action.setChecked(False)

            # Recarrega dados
            self.load_data()
            self.region_label.setText("Visualizando: Geral (Todas)")
            self.populate_table()
            self.update_metrics()
            self.setCursor(Qt.ArrowCursor)
            self.visualizacao_bibtex_widget.carregar_excel()
            self.estatisticas_bibtex_widget.reload_data()
 
        except Exception as e:
            self.setCursor(Qt.ArrowCursor)
            traceback.print_exc()
            QMessageBox.critical(self, "Erro", f"Ocorreu um erro ao atualizar os dados:\n{e}")


    def converter_para_link_csv(self, url):
        match = re.search(r"/d/([a-zA-Z0-9_-]+)", url)
        gid_match = re.search(r"gid=([0-9]+)", url)
        if match:
            sheet_id = match.group(1)
            gid = gid_match.group(1) if gid_match else "0"
            return f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"
        else:
            return url  # retorna como está se não for reconhecido
        

    def exportar_metricas_detalhepara_csv(self):
        texto = self.metricas_textedit.toPlainText().strip()
        if not texto:
            QMessageBox.warning(self, "Aviso", "Não há métricas para exportar.")
            return
        try:
            exportar_metricas_texto_para_csv(texto)

            desktop = QStandardPaths.writableLocation(QStandardPaths.DesktopLocation)

            # Caminho para a pasta PYMT dentro da área de trabalho
            pasta_pymt = os.path.join(desktop, "PYMT")

            # Cria a pasta se ela não existir
            os.makedirs(pasta_pymt, exist_ok=True)

            # Caminho completo do arquivo a ser salvo
            
            file_path = os.path.join(pasta_pymt, "metricas_detalhe.csv")

            QMessageBox.information(self, "Sucesso", f"Métricas exportadas com sucesso para:\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao salvar o CSV:\n{e}")
    def exportar_metricas_para_csv(self):
        texto = self.metricas_textedit.toPlainText().strip()
        if not texto:
            QMessageBox.warning(self, "Aviso", "Não há métricas para exportar.")
            return
        try:
            exportar_metricas_texto_para_csv(texto)

            desktop = QStandardPaths.writableLocation(QStandardPaths.DesktopLocation)

            # Caminho para a pasta PYMT dentro da área de trabalho
            pasta_pymt = os.path.join(desktop, "PYMT")

            # Cria a pasta se ela não existir
            os.makedirs(pasta_pymt, exist_ok=True)

            # Caminho completo do arquivo a ser salvo
            
            file_path = os.path.join(pasta_pymt, "metricas.csv")

            QMessageBox.information(self, "Sucesso", f"Métricas exportadas com sucesso para:\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao salvar o CSV:\n{e}")
    def load_data(self):
        try:
            df = pd.read_csv(self.sheet_csv_url)
            

            df = self.expand_ref_column(df)
            df = self.expand_authors_column(df)
            df = self.expand_affiliations_column(df)
            if "Num de Ref" in df.columns:
                def format_num_de_ref(val):
                    try:
                        return str(int(float(val)))
                    except (ValueError, TypeError):
                        return str(val)
                df["Num de Ref"] = df["Num de Ref"].apply(format_num_de_ref)
            self.dataframe = df
            self.update_region_menu()
            self.update_author_menu()
            self.populate_table()
            self.update_metrics()
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao carregar dados:\n{e}")  
    def expand_ref_column(self, df):
        if "Ref" not in df.columns:
            return df
        new_rows = []
        for _, row in df.iterrows():
            ref_cell = row["Ref"]
            if pd.isna(ref_cell):
                new_rows.append(row)
                continue
            parts = str(ref_cell).split("\n\n")
            if len(parts) == 1:
                new_rows.append(row)
            else:
                first_part = parts[0].strip()
                new_row = row.copy()
                new_row["Ref"] = first_part
                new_rows.append(new_row)
                for part in parts[1:]:
                    new_row = pd.Series(index=df.columns)
                    for col in df.columns:
                        new_row[col] = part.strip() if col == "Ref" else ""
                    new_rows.append(new_row)
        return pd.DataFrame(new_rows).reset_index(drop=True)
    def expand_affiliations_column(self, df):
        if "Afiliation" not in df.columns:
            return df
        new_rows = []
        for _, row in df.iterrows():
            afil_cell = row["Afiliation"]
            if pd.isna(afil_cell):
                new_rows.append(row)
                continue
            afiliacoes = [a.strip() for a in re.split(r'\.\s*', str(afil_cell)) if a.strip()]
            if len(afiliacoes) == 1:
                new_rows.append(row)
            else:
                for afil in afiliacoes:
                    new_row = row.copy()
                    new_row["Afiliation"] = afil
                    new_rows.append(new_row)
        return pd.DataFrame(new_rows).reset_index(drop=True)
    def expand_authors_column(self, df):
        if "Autores" not in df.columns or "Afiliation" not in df.columns:
            return df
        new_rows = []
        for _, row in df.iterrows():
            autores_cell = row["Autores"]
            afil_cell = row["Afiliation"]
            if pd.isna(autores_cell):
                new_rows.append(row)
                continue
            autores = [a.strip() for a in re.split(r',\s*(?=[A-Z])', str(autores_cell)) if a.strip()]
            afiliacoes = [a.strip() for a in re.split(r'\.\s+|\.$', str(afil_cell).replace("\n", " ")) if a.strip()]
            if len(afiliacoes) == 1 and len(autores) > 1:
                afiliacoes = afiliacoes * len(autores)
            if len(autores) != len(afiliacoes):
                print(f"[⚠️ Aviso] {len(autores)} autores e {len(afiliacoes)} afiliações não coincidem.")
                print("-> Autores:", autores)
                print("-> Afiliacoes:", afiliacoes)
            max_len = max(len(autores), len(afiliacoes))
            for i in range(max_len):
                new_row = row.copy()
                autor = autores[i] if i < len(autores) else ""
                afil = afiliacoes[i] if i < len(afiliacoes) else ""
                new_row["Autores"] = autor
                new_row["Afiliation"] = afil
                country = ""
                if afil:
                    match = re.search(r'[,;]\s*([^,;]+)$', afil)
                    if match:
                        country = match.group(1).strip()
                new_row["country"] = country
                new_rows.append(new_row)
        return pd.DataFrame(new_rows).reset_index(drop=True)
    def update_region_menu(self):
        self.region_menu.clear()
        self.selected_region = None
        action_todas = QAction("Todas", self)
        action_todas.triggered.connect(lambda: self.filtrar_por_regiao(None))
        self.region_menu.addAction(action_todas)
        if "Region" not in self.dataframe.columns:
            return
        regioes = sorted(self.dataframe["Region"].dropna().unique())
        for regiao in regioes:
            action = QAction(regiao, self)
            action.triggered.connect(lambda checked, r=regiao: self.filtrar_por_regiao(r))
            self.region_menu.addAction(action)
    def update_author_menu(self):
        self.author_menu.clear()
        action_todos = QAction("Todos", self)
        action_todos.triggered.connect(self.clear_autor_filters)
        self.author_menu.addAction(action_todos)
        self.order_by_frequency = getattr(self, 'order_by_frequency', True) 
        if self.order_by_frequency:
            order_label = "Ordenar por: Alfabética"
        else:
            order_label = "Ordenar por: Aparições"
        action_toggle_sort = QAction(order_label, self)
        action_toggle_sort.triggered.connect(self.toggle_autor_ordering)
        self.author_menu.addAction(action_toggle_sort)
        self.author_menu.addSeparator()
        if "Autores" not in self.dataframe.columns:
            return
        contagem_autores = self.dataframe["Autores"].dropna().value_counts()
        if self.order_by_frequency:
            contagem_autores = contagem_autores  # Já está por contagem decrescente
        else:
            contagem_autores = contagem_autores.sort_index(key=lambda x: x.str.lower())
        self.autor_actions = []
        for autor, count in contagem_autores.items():
            action = QAction(f"{autor} ({count})", self)
            action.setCheckable(True)
            action.toggled.connect(self.autor_selection_changed)
            action.setData(autor)
            self.author_menu.addAction(action)
            self.autor_actions.append(action)


    def salvar_detalhes_em_csv(self):
        texto = self.metricas_detalhes_textedit.toPlainText()
        desktop_path = QStandardPaths.writableLocation(QStandardPaths.DesktopLocation)
        pasta_pymt = os.path.join(desktop_path, "PYMT")
        if not os.path.exists(pasta_pymt):
            os.makedirs(pasta_pymt)
        arquivo_csv = os.path.join(pasta_pymt, "detalhes_salvos.csv")

        linhas = texto.splitlines()

        dados = {
            "Tipos de publicação": [],
            "Tipos de delineamento": [],
            "Países mais frequentes": [],
            "Regiões mais frequentes": [],
            "Autores mais frequentes": []
        }

        categoria_atual = None

        padrao_generico = re.compile(r'-\s*(.+?):\s*(\d+)\s*\(([\d\.]+)%\)')
        padrao_pais = re.compile(r'-\s*(.+?):\s*(\d+)\s*referências\s*\(([\d\.]+)%\)')
        padrao_regiao = re.compile(r'-\s*(.+?):\s*(\d+)')
        padrao_autor = re.compile(r'-\s*(.+?):\s*(\d+)\s*aparições')

        for linha in linhas:
            linha = linha.strip()
            if linha.startswith("📚"):
                categoria_atual = "Tipos de publicação"
            elif linha.startswith("🧪"):
                categoria_atual = "Tipos de delineamento"
            elif linha.startswith("🌍"):
                categoria_atual = "Países mais frequentes"
            elif linha.startswith("🗺️"):
                categoria_atual = "Regiões mais frequentes"
            elif linha.startswith("👤"):
                categoria_atual = "Autores mais frequentes"
            elif linha.startswith("-") and categoria_atual:
                if categoria_atual in ("Tipos de publicação", "Tipos de delineamento"):
                    m = padrao_generico.match(linha)
                    if m:
                        nome, valor, perc = m.groups()
                        dados[categoria_atual].append([nome, int(valor), float(perc)])
                elif categoria_atual == "Países mais frequentes":
                    m = padrao_pais.match(linha)
                    if m:
                        nome, valor, perc = m.groups()
                        dados[categoria_atual].append([nome, int(valor), float(perc)])
                elif categoria_atual == "Regiões mais frequentes":
                    m = padrao_regiao.match(linha)
                    if m:
                        nome, valor = m.groups()
                        dados[categoria_atual].append([nome, int(valor), 0.0])
                elif categoria_atual == "Autores mais frequentes":
                    m = padrao_autor.match(linha)
                    if m:
                        nome, valor = m.groups()
                        dados[categoria_atual].append([nome, int(valor), 0.0])

        # Calcular % para Regiões
        regioes = dados["Regiões mais frequentes"]
        total_regioes = sum(item[1] for item in regioes)
        if total_regioes > 0:
            for item in regioes:
                item[2] = round(item[1] / total_regioes * 100, 1)

        # Ajustar strings para CSV
        for categoria in dados:
            for i, item in enumerate(dados[categoria]):
                nome, valor, perc = item
                perc_str = f"{perc}%" if perc > 0 else ""
                dados[categoria][i] = [nome, str(valor), perc_str]

        # Alinhar colunas
        max_len = max(len(lst) for lst in dados.values())
        for key in dados:
            while len(dados[key]) < max_len:
                dados[key].append(["", "", ""])

        try:
            with open(arquivo_csv, mode='w', newline='', encoding='utf-8') as file:
                writer = csv.writer(file)
                header = []
                for categoria in dados.keys():
                    header.extend([f"{categoria}", "Valor", "% do total"])
                writer.writerow(header)

                for i in range(max_len):
                    row = []
                    for categoria in dados.keys():
                        row.extend(dados[categoria][i])
                    writer.writerow(row)

            QMessageBox.information(self, "Sucesso", f"Arquivo salvo em:\n{arquivo_csv}")
        except Exception as e:
            QMessageBox.critical(self, "Erro ao salvar", f"Erro: {str(e)}")



    def toggle_autor_ordering(self):
        self.order_by_frequency = not self.order_by_frequency
        self.update_author_menu()
    def clear_autor_filters(self):
        for action in getattr(self, "autor_actions", []):
            action.setChecked(False)
        self.filtrar_por_autores([])
    def autor_selection_changed(self, checked):
        autores_selecionados = [
            action.data() for action in getattr(self, "autor_actions", [])
            if action.isChecked()]
        self.filtrar_por_autores(autores_selecionados)
    def filtrar_por_autores(self, autores):
        if not autores:
            self.region_label.setText("Visualizando: Geral (Todas)")
            self.populate_table()
         
            return
        df = self.dataframe.copy()
        df_filtrado = df[df["Autores"].isin(autores)]
        self.region_label.setText(f"Visualizando: Autor(es) - {', '.join(autores)}")
        self.populate_table_custom(df_filtrado)
       

    def filtrar_por_regiao(self, regiao):
        self.selected_region = regiao
        self.region_label.setText(f"Visualizando: {regiao if regiao else 'Geral (Todas)'}")
        self.populate_table()
       
    def populate_table(self):
        df = self.dataframe
        if self.selected_region:
            df = df[df["Region"] == self.selected_region]
        self.populate_table_custom(df)
    def populate_table_custom(self, df):
        self.table.clear()
        self.table.setRowCount(len(df))
        self.table.setColumnCount(len(df.columns))
        self.table.setHorizontalHeaderLabels(df.columns)
        font_bold = QFont()
        font_bold.setBold(True)
        region_colors = {
            "Africa": QColor(240, 230, 140),
            "Asia": QColor(173, 216, 230),
            "Australia and New Zeland": QColor(255, 235, 205),
            "Canada and EUA": QColor(200, 200, 255),
            "Europe": QColor(221, 160, 221),
            "Latin America": QColor(152, 251, 152),
            "Eastern Mediterranean": QColor(255, 182, 193),
        }
        for row in range(len(df)):
            region = df.iloc[row].get("Region", "").strip() if "Region" in df.columns else ""
            bg_color = region_colors.get(region, QColor(255, 255, 255))
            for col, column_name in enumerate(df.columns):
                value = str(df.iat[row, col])
                item = QTableWidgetItem(value)
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                if value.strip() == "":
                    item.setBackground(QColor(220, 220, 220))
                    item.setForeground(QColor(0, 0, 0))
                    item.setTextAlignment(Qt.AlignCenter)
                else:
                    item.setBackground(bg_color)
                    item.setForeground(QColor(0, 0, 0))

                if column_name.lower() == "num":
                    item.setFont(font_bold)
                    item.setTextAlignment(Qt.AlignCenter)
                self.table.setItem(row, col, item)
        self.clear_duplicate_cells([
    "Num", "Region", "country", "Autores", "Titulo", "Ref", "Afiliation",
    "Abstract", "Num de Ref", "IA abstract 100 palavras", "IA keywords"
])
        self.update_spans()
        self.color_region_blocks()
        self.update_metrics()
    def clear_duplicate_cells(self, columns):
        header_labels = [self.table.horizontalHeaderItem(i).text() for i in range(self.table.columnCount())]
        col_indices = [header_labels.index(col) for col in columns if col in header_labels]
        for col in col_indices:
            last_value = None
            for row in range(self.table.rowCount()):
                item = self.table.item(row, col)
                if item:
                    current_value = item.text().strip()
                    if current_value == last_value:
                        item.setText("")
                    else:
                        last_value = current_value
    def update_spans(self):
        self.clear_spans()
        header_labels = [self.table.horizontalHeaderItem(i).text() for i in range(self.table.columnCount())]
        columns_conditional_merge = [
            "Num", "Region", "country", "Autores", "Titulo", "Ref", "Afiliation", "Abstract",
            "Num de Ref", "IA abstract 100 palavras", "IA keywords"]
        for col_name in columns_conditional_merge:
            if col_name not in header_labels:
                continue
            col = header_labels.index(col_name)
            row = 0
            while row < self.table.rowCount():
                item = self.table.item(row, col)
                if item is None or item.text().strip() == "":
                    row += 1
                    continue
                start_row = row
                span_count = 1
                next_row = row + 1
                while next_row < self.table.rowCount():
                    next_item = self.table.item(next_row, col)
                    if next_item and next_item.text().strip() == "":
                        span_count += 1
                        next_row += 1
                    else:
                        break
                if span_count > 1:
                    self.table.setSpan(start_row, col, span_count, 1)
                row = next_row
    def clear_spans(self):
        for r in range(self.table.rowCount()):
            for c in range(self.table.columnCount()):
                if self.table.rowSpan(r, c) > 1 or self.table.columnSpan(r, c) > 1:
                    self.table.setSpan(r, c, 1, 1)
    def color_region_blocks(self):
        try:
            region_col_index = self.dataframe.columns.get_loc("Region")
        except KeyError:
            return
        block_color_1 = QColor(255, 255, 255)
        block_color_2 = QColor(245, 245, 245)
        last_value = None
        block_num = 0
        for row in range(self.table.rowCount()):
            item = self.table.item(row, region_col_index)
            if item is None:
                continue
            current_value = item.text().strip()
            if current_value != last_value:
                block_num += 1
                last_value = current_value
            block_color = block_color_1 if (block_num % 2) == 1 else block_color_2
            for col in range(self.table.columnCount()):
                cell = self.table.item(row, col)
                if cell:
                    bg = cell.background().color()
                    blended = QColor(
                        (bg.red() + block_color.red()) // 2,
                        (bg.green() + block_color.green()) // 2,
                        (bg.blue() + block_color.blue()) // 2,)
                    cell.setBackground(blended)                
    def update_metrics(self):
        try:
            # Usa a função que calcula métricas + estatísticas numéricas detalhadas
            metricas_str = calcular_metricas(self.dataframe)
            self.metricas_textedit.setPlainText(metricas_str)
            self.update_detalhes()
            self.update_entrada_bibtex()
            self.update_visualizacao_bibtex()
            self.update_estatisticas_bibtex()

        except Exception as e:
            self.metrics_label.setText("Erro ao calcular métricas.")
            self.metricas_textedit.setPlainText(f"Erro ao calcular métricas:\n{e}")
    def update_detalhes(self):
        try:
            detalhes_str = calcular_metricas_detalhadas(self.dataframe)
            self.metricas_detalhes_textedit.setPlainText(detalhes_str)
        except Exception as e:
            self.metricas_detalhes_textedit.setPlainText(f"Erro ao calcular métricas detalhadas:\n{e}")

    def update_entrada_bibtex(self):
        try:
            texto = processar_entrada_bibtex(self.dataframe)
            self.entrada_bibtex_textedit.setPlainText(texto)
        except Exception as e:
            self.entrada_bibtex_textedit.setPlainText(f"Erro ao atualizar Entrada Bibtex:\n{e}")

    def update_visualizacao_bibtex(self):
        try:
            texto_formatado = processar_visualizacao_formatada(self.dataframe)
            self.visualizacao_bibtex_textedit.setPlainText(texto_formatado)
        except Exception as e:
            self.visualizacao_bibtex_textedit.setPlainText(f"Erro ao atualizar Visualização Formatada:\n{e}")

    def update_estatisticas_bibtex(self):
        try:
            estatisticas = processar_estatisticas_bibtex(self.dataframe)
            self.estatisticas_bibtex_textedit.setPlainText(estatisticas)
        except Exception as e:
            self.estatisticas_bibtex_textedit.setPlainText(f"Erro ao atualizar Estatísticas Bibtex:\n{e}")

    def export_to_csv(self):
        headers = [self.table.horizontalHeaderItem(i).text() for i in range(self.table.columnCount())]
        if not headers:
            QMessageBox.warning(self, "Aviso", "Nada para exportar.")
            return
        rows = []
        for row in range(self.table.rowCount()):
            row_data = []
            for col in range(self.table.columnCount()):
                item = self.table.item(row, col)
                row_data.append(item.text() if item else "")
            rows.append(row_data)
        df_export = pd.DataFrame(rows, columns=headers)
        desktop = QStandardPaths.writableLocation(QStandardPaths.DesktopLocation)

        # Caminho para a pasta PYMT dentro da área de trabalho
        pasta_pymt = os.path.join(desktop, "PYMT")

        # Cria a pasta se ela não existir
        os.makedirs(pasta_pymt, exist_ok=True)

        # Caminho completo do arquivo a ser salvo
        file_path = os.path.join(pasta_pymt,  "exported_table.csv")

        try:
            df_export.to_csv(file_path, index=False)
            QMessageBox.information(self, "Sucesso", f"Arquivo salvo com sucesso em:\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao salvar arquivo:\n{e}")



if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = GoogleSheetsViewer()
    window.show()
    sys.exit(app.exec_())
