import sys, re, os, requests
import pandas as pd
from collections import Counter
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QPushButton, QTextBrowser, QLabel, QFileDialog, 
                             QTabWidget, QTableWidget, QTableWidgetItem, 
                             QHeaderView, QHBoxLayout, QFrame, QScrollArea, QInputDialog, QLineEdit, QComboBox,QSlider,
                             QSplitter, QMessageBox, QDialog,  QTextEdit)

from PyQt5.QtCore import  Qt, QTimer, QUrl
from PyQt5.QtGui import QColor, QFont
from PyQt5.QtWebEngineWidgets import QWebEngineView, QWebEngineSettings 
from workers import DeepExtractionWorker, BibtexWorker, DonutChart, BarChart,EditableDetailDialog


import networkx as nx
from sklearn.feature_extraction.text import CountVectorizer
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import matplotlib.pyplot as plt

class GraphConfigDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Configurações da Rede EGA")
        self.setMinimumWidth(350)
        self.setStyleSheet("background: #111; color: #EEE;")
        
        layout = QVBoxLayout(self)

        # Filtro de Conexão (Threshold)
        layout.addWidget(QLabel("Força da Conexão (Threshold):"))
        self.threshold_slider = QSlider(Qt.Horizontal)
        self.threshold_slider.setRange(5, 80) # 0.05 a 0.80
        self.threshold_slider.setValue(25)
        layout.addWidget(self.threshold_slider)

        # Cor da Fonte
        layout.addWidget(QLabel("Cor da Fonte:"))
        self.combo_font_color = QComboBox()
        self.combo_font_color.addItems(["White", "Cyan", "Yellow", "LightGray"])
        layout.addWidget(self.combo_font_color)

        # Estilo do Layout (NetworkX)
        layout.addWidget(QLabel("Estilo do Layout:"))
        self.combo_layout = QComboBox()
        self.combo_layout.addItems(["Spring (Orgânico)", "Circular", "Shell", "Spectral"])
        layout.addWidget(self.combo_layout)

        # Botão Confirmar
        self.btn_ok = QPushButton("GERAR GRÁFICO")
        self.btn_ok.setStyleSheet("background: #50F; color: white; padding: 10px;")
        self.btn_ok.clicked.connect(self.accept)
        layout.addWidget(self.btn_ok)

    def get_values(self):
        return {
            "threshold": self.threshold_slider.value() / 100,
            "font_color": self.combo_font_color.currentText().lower(),
            "layout": self.combo_layout.currentText()
        }
class KeywordNetworkWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(0, 0, 0, 0)
        
        # Configuração do Matplotlib estilo Dark
        plt.style.use('dark_background')
        self.figure = Figure(figsize=(8, 6), facecolor='#050505')
        self.canvas = FigureCanvas(self.figure)
        self.canvas.setStyleSheet("background-color: #050505;")
        self.ax = self.figure.add_subplot(111)
        self.ax.set_facecolor('#050505')
        
        self.layout.addWidget(self.canvas)



    def get_canonical_author(self, name):
        norm = self.normalize_author_name(name)
        if not norm: return None
        
        # Remove vírgulas e separa as partes do nome
        parts = norm.replace(",", "").split()
        # Ordena as partes alfabeticamente: ['JOAO', 'SILVA'] -> 'JOAO_SILVA'
        # Isso faz com que 'Silva, Joao' e 'Joao Silva' gerem o mesmo ID
        parts.sort()
        return "_".join(parts)
    def generate_graph(self, text_list, min_freq=2, top_n=50, threshold=0.3, font_color='white', layout_type='Spring (Orgânico)', node_color_theme='Cyberpunk'):
            """
            Gera o grafo de clusters (EGA style) com parâmetros customizados.
            """
            self.ax.clear()
            
            if not text_list:
                self.ax.text(0.5, 0.5, "Sem dados para gerar grafo", 
                            ha='center', color='#555', fontsize=14)
                self.canvas.draw()
                return

            # 1. Vetorização e Matriz de Co-ocorrência
            my_stops = list(CountVectorizer(stop_words='english').get_stop_words())
            my_stops.extend(['study', 'results', 'using', 'based', 'analysis', 'paper', 'proposed', 'data', 'research'])

            vectorizer = CountVectorizer(stop_words=my_stops, min_df=min_freq, max_features=top_n)
            try:
                X = vectorizer.fit_transform(text_list)
            except ValueError:
                return

            xc = (X.T * X)
            xc.setdiag(0)
            
            names = vectorizer.get_feature_names_out()
            df_co = pd.DataFrame(data=xc.toarray(), columns=names, index=names)

            # 2. Criação do Grafo
            G = nx.Graph()
            max_val = df_co.max().max() if not df_co.empty else 1
            
            for i in range(len(names)):
                for j in range(i + 1, len(names)):
                    weight = df_co.iloc[i, j]
                    norm_weight = weight / max_val 
                    if norm_weight > threshold:
                        G.add_edge(names[i], names[j], weight=norm_weight)

            G.remove_nodes_from(list(nx.isolates(G)))

            if G.number_of_nodes() == 0:
                self.ax.text(0.5, 0.5, "Nenhuma conexão forte encontrada.\nReduza o Threshold.", 
                            ha='center', color='#555')
                self.canvas.draw()
                return

            # 3. Lógica de Cores dos Clusters
            from networkx.algorithms import community
            try:
                communities = community.greedy_modularity_communities(G)
                node_groups = {}
                
                # Definindo paletas baseadas no tema
                if node_color_theme == 'Pastel':
                    colors = ['#FFB7B2', '#FFDAC1', '#E2F0CB', '#B5EAD7', '#C7CEEA']
                elif node_color_theme == 'Ocean':
                    colors = ['#0077B6', '#00B4D8', '#90E0EF', '#ADE8F4', '#CAF0F8']
                else: # Default Cyberpunk
                    colors = ['#00D2FF', '#7B2FF7', '#FF3366', '#FFD700', '#33FF77']
                
                for i, com in enumerate(communities):
                    for node in com:
                        node_groups[node] = i
                
                color_list = [colors[node_groups[n] % len(colors)] for n in G.nodes()]
            except:
                color_list = '#00D2FF'

            # 4. Seleção do Layout (vindo do Popup)
            if "Circular" in layout_type:
                pos = nx.circular_layout(G)
            elif "Shell" in layout_type:
                pos = nx.shell_layout(G)
            elif "Spectral" in layout_type:
                pos = nx.spectral_layout(G)
            else: # Spring é o padrão
                pos = nx.spring_layout(G, k=0.5, iterations=50, seed=42)
            
            # 5. Desenho
            # Arestas
            edges = G.edges(data=True)
            weights = [d['weight'] * 4 for u, v, d in edges] 
            nx.draw_networkx_edges(G, pos, ax=self.ax, width=weights, alpha=0.3, edge_color='#555')

            # Nós
            d = dict(G.degree)
            node_sizes = [v * 150 + 200 for v in d.values()]
            nx.draw_networkx_nodes(G, pos, ax=self.ax, node_size=node_sizes, 
                                node_color=color_list, alpha=0.85, 
                                linewidths=1.5, edgecolors='#111')


            labels = {node: node.replace("_", " ").title() for node in G.nodes()}

            nx.draw_networkx_labels(G, pos, labels=labels, ax=self.ax, font_size=10, 
                                    font_color=font_color, font_family='Arial', 
                                    font_weight='bold')

            self.ax.axis('off')
            self.figure.tight_layout(pad=0)
            self.canvas.draw()


class BiblioApp(QMainWindow):
    def __init__(self):
        super().__init__(); 
        self.live_data = [] # Lista acumulativa
        self.is_processing = False
        self.current_progress = (0, 0)  # (atual, total)
        self.spinner_chars = ["⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏"]
        self.spin_idx = 0
        self.initUI()
        self.spinner_timer = QTimer()
        self.spinner_timer.timeout.connect(self.update_progress_label)


    def initUI(self):
            self.app_name = "BiblioTech - Multi-Source Accumulator"
            self.setWindowTitle("BiblioTech - Multi-Source Accumulator"); self.resize(1600, 1000)
            self.setStyleSheet("""
                QMainWindow, QWidget { background: #050505; color: #FFF; font-size: 14px; font-family: 'Segoe UI', Arial; }
                QTableWidget { background: #020202; gridline-color: #111; color: #CCC; border: none; font-size: 12px; }
                QHeaderView::section { background: #0A0A0A; color: #00D2FF; font-weight: bold; padding: 8px; border: none; font-size: 12px; }
                               
                               
/* Botão Padrão: Ajustável */
QPushButton { 
    background: #111; 
    border: 1px solid #333; 
    color: #BBB; 
    padding: 5px 15px;      /* Padding lateral mantém o respiro, mas permite encolher */
    font-weight: bold; 
    font-size: 13px; 
    min-width: 0px;         /* REMOVIDO os 100px para permitir ajuste ao texto */
    border-radius: 6px; 
}

QPushButton:hover { 
    background: #222; 
    color: #FFF; 
    border-color: #555; 
}

/* Botões de Canto (Tabs): Compactos */
QPushButton#CornerTab { 
    background: #0A0A0A; 
    color: #888; 
    border: 1px solid #1A1A1A; 
    border-bottom: none;  
    border-top-left-radius: 6px;  
    border-top-right-radius: 6px;  
    padding: 8px 12px;      /* Reduzido de 12px 20px para não ficar "inchado" */
    font-size: 11px; 
    letter-spacing: 0.5px;
    text-transform: uppercase;      
    margin-top: 5px;       
    min-width: 20px;        /* Permite que o botão seja tão pequeno quanto o ícone ↻ */
}

QPushButton#CornerTab:hover {   
    background: #151515;   
    color: #00D2FF;    
    border-color: #333;   
}

/* Botão de Ação: Mantém um destaque, mas sem largura exagerada */
QPushButton#ActionBtn { 
    background: #003344; 
    color: #00D2FF; 
    border: 1px solid #005577; 
    padding: 8px 20px; 
    min-width: 80px; 
}

QPushButton#ActionBtn:hover { 
    background: #004455; 
}
                QTabWidget::pane { border-top: 1px solid #151515; } 
                QTabBar::tab { background: #0A0A0A; padding: 12px 25px; color: #666; font-weight: bold; font-size: 14px; border-top-left-radius: 6px; border-top-right-radius: 6px; }
                QTabBar::tab:selected { color: #FFF; border-bottom: 2px solid #00D2FF; background: #0F0F0F; }
                QSplitter::handle { background: #111; }
            """)

            main_container = QWidget()
            main_layout = QVBoxLayout(main_container)
            main_layout.setContentsMargins(0, 0, 0, 0)
            main_layout.setSpacing(0)

            self.tabs = QTabWidget()
            self.tabs.setDocumentMode(True)
            self.left_widget = QWidget()
            self.left_layout = QHBoxLayout(self.left_widget)
            self.left_layout.setContentsMargins(10, 5, 10, 5)
            
            self.btn_load = QPushButton("IMPORT pdf bib")
            self.btn_load.setObjectName("ActionBtn")
            self.btn_load.clicked.connect(self.open_pdf)
            
            self.lbl_count = QLabel("0 Refs")
            self.lbl_count.setStyleSheet("color: #00D2FF; font-weight: bold; margin-left: 10px;")
            
            self.left_layout.addWidget(self.btn_load)
            self.left_layout.addWidget(self.lbl_count)
            self.tabs.setCornerWidget(self.left_widget, Qt.TopLeftCorner)

            self.corner_db_widget = QWidget()
            self.corner_db_layout = QHBoxLayout(self.corner_db_widget)
            self.corner_db_layout.setContentsMargins(0, 0, 5, 0) 
            self.corner_db_layout.setSpacing(2) # Espaço mínimo entre os botões

            self.corner_dash_widget = QWidget()
            self.corner_dash_layout = QHBoxLayout(self.corner_dash_widget)
            self.corner_dash_layout.setContentsMargins(0, 0, 5, 0)
            self.corner_dash_layout.setSpacing(2)

            self.btn_sync_table = QPushButton("↻")
            self.btn_sync_table.setObjectName("CornerTab")
            self.btn_sync_table.clicked.connect(self.update_table_view)

            self.btn_bib = QPushButton(".bib")

            self.btn_apa = QPushButton("APA")
            self.btn_abnt = QPushButton("ABNT")

            self.btn_apa.setObjectName("CornerTab")
            self.btn_abnt.setObjectName("CornerTab")

            self.btn_apa.clicked.connect(self.export_apa)
            self.btn_abnt.clicked.connect(self.export_abnt)


            self.btn_bib.setObjectName("CornerTab")
            self.btn_bib.clicked.connect(self.export_bibtex)

            
            self.corner_db_layout.addWidget(self.btn_bib)
            self.corner_db_layout.addWidget(self.btn_abnt)
            self.corner_db_layout.addWidget(self.btn_apa)
            self.corner_db_layout.addWidget(self.btn_sync_table)
            

            self.corner_db_widget.hide() 

            # -- GRUPO B: DASHBOARD BUTTONS --
            self.corner_dash_widget = QWidget()
            self.corner_dash_layout = QHBoxLayout(self.corner_dash_widget)
            self.corner_dash_layout.setContentsMargins(0, 0, 10, 0)
            self.corner_dash_layout.setSpacing(5)

            self.btn_sync_dash = QPushButton("↻")
            self.btn_sync_dash.setObjectName("CornerTab")
            self.btn_sync_dash.clicked.connect(self.update_dashboard_view)

            self.btn_xlsx = QPushButton("EXPORT .xlsx")
            self.btn_xlsx.setObjectName("CornerTab")
            self.btn_xlsx.clicked.connect(self.export_excel)
            self.corner_dash_layout.addWidget(self.btn_xlsx)
            self.corner_dash_layout.addWidget(self.btn_sync_dash)
            
            self.corner_dash_widget.hide()
# --- GRUPO C: KEYWORDS BUTTONS (Canto Direito) ---
            self.corner_kw_widget = QWidget()
            self.corner_kw_layout = QHBoxLayout(self.corner_kw_widget)
            self.corner_kw_layout.setContentsMargins(0, 0, 10, 0)
            self.corner_kw_layout.setSpacing(5)


            self.corner_kw_widget = QWidget()
            self.corner_kw_layout = QHBoxLayout(self.corner_kw_widget)
            self.corner_kw_layout.setContentsMargins(0, 0, 10, 0)
            self.corner_kw_layout.setSpacing(5)

            # 1. Botão de Gerar Rede com estilo Roxo Neon
            self.btn_gen_graph = QPushButton("GERAR REDE (EGA)")
            self.btn_gen_graph.setObjectName("CornerTab") 
            # ESTILO ROXO: Mantém a base do CornerTab mas muda cor e borda
            self.btn_gen_graph.setStyleSheet("""
                QPushButton#CornerTab { 
                    color: #D0F; 
                    border: 1px solid #50F; 
                    background: #1A0033;
                }
                QPushButton#CornerTab:hover { 
                    background: #2A0044; 
                    color: #E0F; 
                    border-color: #70F; 
                }
            """)
            self.btn_gen_graph.clicked.connect(self.run_network_analysis)

            self.btn_extract_kw = QPushButton("↻")
            self.btn_extract_kw.setObjectName("CornerTab")
            self.btn_extract_kw.clicked.connect(self.run_keyword_extraction)

            self.btn_export_kw_excel = QPushButton("EXPORT .xlsx")
            self.btn_export_kw_excel.setObjectName("CornerTab")
            self.btn_export_kw_excel.clicked.connect(self.export_keywords_excel)
            self.corner_kw_layout.addWidget(self.btn_gen_graph)
            self.corner_kw_layout.addWidget(self.btn_export_kw_excel)
            self.corner_kw_layout.addWidget(self.btn_extract_kw)
            
            self.corner_kw_widget.hide()

            # --- GRUPO D: AUTHORS BUTTONS (Canto Direito) ---
            self.corner_auth_widget = QWidget()
            self.corner_auth_layout = QHBoxLayout(self.corner_auth_widget)
            self.corner_auth_layout.setContentsMargins(0, 0, 10, 0)
            self.corner_auth_layout.setSpacing(5)

            self.btn_gen_auth_graph = QPushButton("GERAR REDE (AUTORES)")
            self.btn_gen_auth_graph.setObjectName("CornerTab") 
            self.btn_gen_auth_graph.setStyleSheet("""
                QPushButton#CornerTab { color: #00D2FF; border: 1px solid #005577; background: #001122; }
                QPushButton#CornerTab:hover { background: #002233; color: #FFF; border-color: #00D2FF; }
            """)
            self.btn_gen_auth_graph.clicked.connect(self.run_author_network_analysis)

            self.btn_extract_auth = QPushButton("↻")
            self.btn_extract_auth.setObjectName("CornerTab")
            self.btn_extract_auth.setFixedWidth(35)
            self.btn_extract_auth.clicked.connect(self.run_author_extraction)

            self.corner_auth_layout.addWidget(self.btn_gen_auth_graph)
            self.corner_auth_layout.addWidget(self.btn_extract_auth)
            self.corner_auth_widget.hide()



            # --- ABA 1: LOG & PDF ---
            self.tab1_widget = QWidget()
            self.tab1_layout = QVBoxLayout(self.tab1_widget)
            self.tab1_layout.setContentsMargins(0, 0, 0, 0)
            
            self.splitter = QSplitter(Qt.Horizontal)
            self.tab_scan = QTextBrowser() 
            self.tab_scan.setOpenExternalLinks(True) 
            self.tab_scan.setStyleSheet("background: #000; border: none; padding: 10px; font-family: 'Segoe UI'; font-size: 13px;")
            
            self.pdf_viewer = QWebEngineView()
            self.pdf_viewer.settings().setAttribute(QWebEngineSettings.PluginsEnabled, True)
            self.pdf_viewer.page().setBackgroundColor(QColor("#050505"))
            self.pdf_viewer.setStyleSheet("background: #050505; border-left: 1px solid #111;")

            self.splitter.addWidget(self.tab_scan)
            self.splitter.addWidget(self.pdf_viewer)
            self.splitter.setSizes([600, 600])
            self.tab1_layout.addWidget(self.splitter)

            # --- ABA 2: TABELA (DATABASE) ---
            self.tab2_widget = QWidget()
            self.tab2_layout = QVBoxLayout(self.tab2_widget)
            self.tab2_layout.setContentsMargins(0, 0, 0, 0)

            self.table = QTableWidget(0, 9)
            self.table.setHorizontalHeaderLabels([
                "FILE", "AUTHORS", "TITLE", "YEAR", "DOI", 
                "JOURNAL", "PUBLISHER", "TYPE", "ABSTRACT"
            ])

            table_font = QFont("Segoe UI", 9)
            self.table.setFont(table_font)
            self.table.horizontalHeader().setFont(table_font)

            self.table.setTextElideMode(Qt.ElideRight)
            self.table.setSelectionBehavior(QTableWidget.SelectRows)

            v_header = self.table.verticalHeader()
            v_header.setDefaultSectionSize(28) 

            header = self.table.horizontalHeader()
            self.table.setColumnWidth(0, 130) # FILE
            self.table.setColumnWidth(3, 50)  # YEAR
            self.table.setColumnWidth(4, 150) # DOI
            header.setSectionResizeMode(2, QHeaderView.Stretch) # TITLE
            header.setSectionResizeMode(8, QHeaderView.Stretch) # ABSTRACT

            self.table.cellClicked.connect(self.handle_cell_click)
            self.tab2_layout.addWidget(self.table)
            
            # --- ABA 3: DASHBOARD ---
            self.tab3_widget = QWidget()
            self.tab3_layout = QVBoxLayout(self.tab3_widget)
            self.tab3_layout.setContentsMargins(0, 0, 0, 0)

            dash_scroll = QScrollArea(); dash_scroll.setWidgetResizable(True); dash_scroll.setStyleSheet("border: none;") 
            self.dash_inner = QWidget(); self.l_dash = QVBoxLayout(self.dash_inner); self.setup_dash_widgets()
            dash_scroll.setWidget(self.dash_inner)
            self.tab3_layout.addWidget(dash_scroll)

# --- ABA 4: KEYWORDS & NETWORK ---
            self.tab4_widget = QWidget()
            self.tab4_layout = QVBoxLayout(self.tab4_widget)
            

            # 2. Splitter Vertical: Tabelas em Cima / Grafo em Baixo
            self.main_kw_splitter = QSplitter(Qt.Vertical)
            
            # --- Parte de Cima: As Tabelas (que você já tinha) ---
            self.tables_container = QWidget()
            self.tables_layout = QHBoxLayout(self.tables_container)
            self.tables_layout.setContentsMargins(0,0,0,0)
            
            self.kw_table_titles = QTableWidget(0, 4)
            self.setup_keyword_table(self.kw_table_titles, "TITLE KEYWORDS")
            self.kw_table_abstracts = QTableWidget(0, 4)
            self.setup_keyword_table(self.kw_table_abstracts, "ABSTRACT KEYWORDS")
            
            self.tables_splitter = QSplitter(Qt.Horizontal)
            self.tables_splitter.addWidget(self.kw_table_titles)
            self.tables_splitter.addWidget(self.kw_table_abstracts)
            
            self.tables_layout.addWidget(self.tables_splitter)
            self.main_kw_splitter.addWidget(self.tables_container)

            # --- Parte de Baixo: O Grafo EGA ---
            self.graph_container = QWidget()
            self.graph_layout = QVBoxLayout(self.graph_container)
            self.graph_layout.setContentsMargins(0,0,0,0)
            
            # Instanciando nossa nova classe
            self.network_widget = KeywordNetworkWidget()
            self.graph_layout.addWidget(self.network_widget)
            
            self.main_kw_splitter.addWidget(self.graph_container)
            self.main_kw_splitter.setSizes([300, 500]) # Define tamanho inicial (tabelas menores, grafo maior)

            self.tab4_layout.addWidget(self.main_kw_splitter)


                        # --- ABA 5: AUTHORS & CO-AUTHORSHIP ---
            self.tab5_widget = QWidget()
            self.tab5_layout = QVBoxLayout(self.tab5_widget)

            self.main_auth_splitter = QSplitter(Qt.Vertical)

            # Tabela de Autores
            self.auth_table = QTableWidget(0, 4)
            self.setup_keyword_table(self.auth_table, "AUTHOR NAME") # Reutiliza o setup de colunas

            # Grafo de Autores
            self.auth_network_widget = KeywordNetworkWidget() # Reutiliza a classe do grafo

            self.main_auth_splitter.addWidget(self.auth_table)
            self.main_auth_splitter.addWidget(self.auth_network_widget)
            self.main_auth_splitter.setSizes([300, 600])

            self.tab5_layout.addWidget(self.main_auth_splitter)


                        
            # ADICIONAR ABAS
            self.tabs.addTab(self.tab1_widget, "TERMINAL") # Index 0
            self.tabs.addTab(self.tab2_widget, "DATABASE")    # Index 1
            self.tabs.addTab(self.tab4_widget, "KEYWORDS")
                        # ADICIONAR ABA NO TABS
            self.tabs.addTab(self.tab5_widget, "AUTHORS")
            self.tabs.addTab(self.tab3_widget, "DASHBOARD")   # Index 2
      

            # CONECTAR SINAL DE TROCA DE ABA
            self.tabs.currentChanged.connect(self.on_tab_changed)

            # --- RODAPÉ DE PROGRESSO ---
            self.progress_label = QLabel("")
            self.progress_label.setStyleSheet("""
                QLabel { color: #00D2FF;  background: #000000; padding: 0px;  font-family: 'Consolas', 'Monaco', monospace;  font-size: 14px;  font-weight: bold;  border-top: 1px solid #1A1A1A;}
            """)
            self.progress_label.setFixedHeight(50) 
            self.progress_label.setAlignment(Qt.AlignCenter)
            self.progress_label.hide()

            main_layout.addWidget(self.tabs)
            main_layout.addWidget(self.progress_label)
            self.setCentralWidget(main_container)


    def setup_keyword_table(self, table, title):
        """Configuração padronizada para as tabelas de keywords."""
        table.setColumnCount(4)
        table.setHorizontalHeaderLabels(["USE", title, "FREQ", "%"])
        table.setStyleSheet("border: none; background: #020202;")
        header = table.horizontalHeader()
        table.setColumnWidth(0, 40)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        table.setColumnWidth(2, 60)
        table.setColumnWidth(3, 60)
        table.verticalHeader().setVisible(False)


    def run_network_analysis(self):
            if not self.live_data:
                QMessageBox.warning(self, "Aviso", "Sem dados carregados.")
                return

            # 1. Abre o Popup de Configuração
            dialog = GraphConfigDialog(self)
            if dialog.exec_() == QDialog.Accepted:
                config = dialog.get_values()
                
                # 2. Coleta os dados (Títulos + Abstracts)
                combined_texts = []
                for row in self.live_data:
                    titulo = str(row[2]) if row[2] else ""
                    abstract = str(row[8]) if row[8] and row[8] != "N/A" else ""
                    combined_texts.append(f"{titulo} {abstract}".lower())

                # 3. Passa as configurações para o gerador
                # Adicionei os novos parâmetros 'font_color' e 'layout_type'
                self.network_widget.generate_graph(
                    combined_texts, 
                    min_freq=2, 
                    top_n=60, 
                    threshold=config["threshold"],
                    font_color=config["font_color"],
                    layout_type=config["layout"]
                )

    def get_top_terms(self, abstract_list, min_freq=2, min_len=5):
            # Blacklist reforçada para eliminar "ruído"
            blacklist = {
                'their', 'about', 'other', 'which', 'there', 'these', 'those', 'through',
                'study', 'research', 'results', 'analysis', 'using', 'based', 'paper',
                'proposed', 'between', 'during', 'present', 'clinical', 'within', 'after',
                'before', 'under', 'while', 'authors', 'copyright', 'rights', 'reserved',
                'university', 'department', 'institute', 'elsevier', 'springer', 'science',
                'associated', 'potential', 'evidence', 'levels', 'related', 'impact'
            }
            
            # Para a porcentagem correta (Presença nos Artigos):
            # Usamos uma lista de sets para saber se a palavra existe no artigo (independente de quantas vezes)
            words_per_article = []
            for text in abstract_list:
                if text and text != "N/A":
                    # Extrai apenas palavras técnicas (mínimo 5 letras por padrão)
                    found = re.findall(r'\b[A-Za-zÀ-ÿ]{' + str(min_len) + r',}\b', text.lower())
                    # Criamos um conjunto único por artigo para não duplicar a contagem no mesmo texto
                    words_per_article.append(set(w for w in found if w not in blacklist))

            # Contagem Global para Frequência (Total de aparições)
            all_words = [w for art in words_per_article for w in art]
            counter = Counter(all_words)
            
            # Filtro final de frequência mínima
            refined = {word: count for word, count in counter.items() if count >= min_freq}
            return Counter(refined)
        
    def on_tab_changed(self, index):
        self.corner_db_widget.hide()
        self.corner_dash_widget.hide()
        self.corner_kw_widget.hide()
        self.corner_auth_widget.hide()

        if index == 1: # DATABASE
            self.tabs.setCornerWidget(self.corner_db_widget, Qt.TopRightCorner)
            self.corner_db_widget.show()
        elif index == 2: # KEYWORDS
            self.tabs.setCornerWidget(self.corner_kw_widget, Qt.TopRightCorner)
            self.corner_kw_widget.show()
        elif index == 3: # AUTHORS (Nova Aba)
            self.tabs.setCornerWidget(self.corner_auth_widget, Qt.TopRightCorner)
            self.corner_auth_widget.show()
        elif index == 4: # DASHBOARD
            self.tabs.setCornerWidget(self.corner_dash_widget, Qt.TopRightCorner)
            self.corner_dash_widget.show()


    def get_active_keywords(self):
        """Retorna apenas as keywords que estão marcadas na tabela."""
        active_data = {}
        for r in range(self.kw_table.rowCount()):
            if self.kw_table.item(r, 0).checkState() == Qt.Checked:
                word = self.kw_table.item(r, 1).text().lower()
                freq = int(self.kw_table.item(r, 2).text())
                active_data[word] = freq
        return Counter(active_data)

    def export_keywords_excel(self):
        # Agora exporta apenas as MARCADAS
        active_counts = self.get_active_keywords()
        if not active_counts:
            QMessageBox.warning(self, "Aviso", "Nenhum termo selecionado!")
            return
            
        path, _ = QFileDialog.getSaveFileName(self, "Exportar Keywords", "", "*.xlsx")
        if path:
            data = [{"Keyword": k.upper(), "Frequência": v} for k, v in active_counts.items()]
            pd.DataFrame(data).to_excel(path, index=False)
            QMessageBox.information(self, "Sucesso", "Exportado com filtros aplicados!")

    def run_keyword_extraction(self):
            if not self.live_data: return

            # 1. Extração de Títulos
            titles_list = [row[2] for row in self.live_data if row[2] and row[2] != "Sem Título"]
            # Usamos min_len=4 para títulos, pois palavras importantes costumam ser menores (ex: 'data')
            counts_titles = self.get_top_terms(titles_list, min_freq=2, min_len=4)
            self.fill_kw_table(self.kw_table_titles, counts_titles)

            # 2. Extração de Abstracts
            abstracts_list = [row[8] for row in self.live_data if row[8] and row[8] != "N/A"]
            counts_abstracts = self.get_top_terms(abstracts_list, min_freq=2, min_len=5)
            self.fill_kw_table(self.kw_table_abstracts, counts_abstracts)

    def fill_kw_table(self, table, counter):
            """Preenche a tabela específica com os dados do counter."""
            table.setRowCount(0)
            total_refs = len(self.live_data)
            top_items = counter.most_common(100)

            for word, freq in top_items:
                row = table.rowCount()
                table.insertRow(row)
                
                chk_item = QTableWidgetItem()
                chk_item.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
                chk_item.setCheckState(Qt.Checked)
                table.setItem(row, 0, chk_item)
                
                table.setItem(row, 1, QTableWidgetItem(word.upper()))
                
                it_freq = QTableWidgetItem(str(freq))
                it_freq.setTextAlignment(Qt.AlignCenter)
                table.setItem(row, 2, it_freq)
                
                perc = (freq / total_refs) * 100
                it_perc = QTableWidgetItem(f"{perc:.1f}%")
                it_perc.setTextAlignment(Qt.AlignCenter)
                it_perc.setForeground(QColor("#00D2FF"))
                table.setItem(row, 3, it_perc)

    def export_keywords_excel(self):
            if self.kw_table.rowCount() == 0:
                QMessageBox.warning(self, "Erro", "Extraia as keywords primeiro.")
                return
                
            path, _ = QFileDialog.getSaveFileName(self, "Exportar Keywords", "", "*.xlsx")
            if path:
                data = []
                for r in range(self.kw_table.rowCount()):
                    data.append({
                        "Keyword": self.kw_table.item(r, 0).text(),
                        "Frequência (Artigos)": self.kw_table.item(r, 1).text(),
                        "Relevância (%)": self.kw_table.item(r, 2).text()
                    })
                pd.DataFrame(data).to_excel(path, index=False)
                QMessageBox.information(self, "Sucesso", "Arquivo Excel gerado com sucesso!")


    def setup_dash_widgets(self):
            # Zeramos as margens do layout principal do dashboard para ocupar a tela toda
            self.l_dash.setContentsMargins(5, 5, 5, 5) 
            self.l_dash.setSpacing(10)
            
            # --- LINHA 1: KPIs (Compactos) ---
            h_kpi = QHBoxLayout()
            h_kpi.setSpacing(10)
            self.k1 = self.create_kpi("TOTAL")
            self.k2 = self.create_kpi("ABSTRACTS") 
            self.k3 = self.create_kpi("MÁX Year")
            self.k4 = self.create_kpi("Sources")
            
            for kpi in [self.k1, self.k2, self.k4, self.k3]:
                h_kpi.addWidget(kpi)
            self.l_dash.addLayout(h_kpi)

            # --- LINHA 2: DONUTS (Gráficos Circulares) ---
            h_donuts = QHBoxLayout()
            h_donuts.setSpacing(10)
            self.c_yrs = DonutChart("Por Ano")
            self.c_tps = DonutChart("Tipos de Documento")
            self.c_pub = DonutChart("Top Publishers")
            
            for donut in [self.c_yrs, self.c_tps, self.c_pub]:
                h_donuts.addWidget(donut)
            self.l_dash.addLayout(h_donuts)

            # --- LINHA 3: BARRAS PRINCIPAIS (Jornais e Autores) ---
            h_bars = QHBoxLayout()
            h_bars.setSpacing(10)
            self.b_journal = BarChart("Journal/Conference", "#FF3366")
            self.b_aut = BarChart("Author", "#7B2FF7")
            
            h_bars.addWidget(self.b_journal)
            h_bars.addWidget(self.b_aut)
            self.l_dash.addLayout(h_bars)

            # --- LINHA 4: ANÁLISE DE TEXTO (Título e Abstracts) ---
            # Colocamos os dois de keywords lado a lado para comparação direta
            h_kw = QHBoxLayout()
            h_kw.setSpacing(10)
            self.b_word_titles = BarChart("Title Keywords", "#00D2FF")
            self.b_word_freq = BarChart("Abstract Keywords", "#33FF77")
            
            h_kw.addWidget(self.b_word_titles)
            h_kw.addWidget(self.b_word_freq)
            self.l_dash.addLayout(h_kw)

            # --- LINHA 5: RANKING DE ARTIGOS (Largura Total) ---
            # Como títulos de artigos são longos, eles precisam de uma linha exclusiva
            # para evitar que a margem esquerda (margin_left) esmague o gráfico
            self.b_top_articles = BarChart("Articles", "#FFD700") 
            self.b_top_articles.margin_left = 400 # Espaço generoso para ler o título do paper
            self.b_top_articles.setMinimumHeight(500) # Mais altura para o ranking principal
            self.l_dash.addWidget(self.b_top_articles)

    def create_kpi(self, t):
        f = QFrame(); f.setStyleSheet("background: #0F0F0F; border: 1px solid #222; border-radius: 12px;")
        f.setFixedHeight(110)
        l = QVBoxLayout(f); 
        title = QLabel(t); title.setStyleSheet("color: #888; font-size: 13px; font-weight: bold;")
        title.setAlignment(Qt.AlignCenter)
        v = QLabel("-"); v.setStyleSheet("color: #FFF; font-size: 32px; font-weight: 900;")
        v.setAlignment(Qt.AlignCenter)
        l.addWidget(title); l.addWidget(v); f.v = v; return f


    def handle_cell_click(self, row, col):
            try:
                item = self.table.item(row, col)
                current_text = item.text().strip() if item else ""
                upper_text = current_text.upper()
                
                # Lista expandida de termos inválidos
                invalid_terms = ["N/A", "N.D.", "SEM TÍTULO", "SEM TITULO", "UNKNOWN", "NONE"]
                
                # Verifica se o conteúdo atual deve ser editado
                is_placeholder = (
                    not current_text or
                    any(term in upper_text for term in invalid_terms) or
                    len(current_text) < 2  # Captura ruídos ou campos vazios
                )

                if is_placeholder:
                    headers = ["Arquivo", "Autores", "Título", "Ano", "DOI", "Jornal", "Publisher", "Tipo", "Abstract"]
                    column_name = headers[col] if col < len(headers) else "Informação"
                    
                    new_text, ok = QInputDialog.getText(
                        self, 'EDIT', 

                        QLineEdit.Normal, ""
                    )
                    
                    if ok and new_text.strip():
                        val = new_text.strip()
                        
                        # Validação específica para a coluna ANO (índice 3)
                        if col == 3 and not val.isdigit():
                            QMessageBox.warning(self, "Error", "Number only.")
                            return

                        # 1. Atualiza a tabela
                        self.table.setItem(row, col, QTableWidgetItem(val))
                        
                        # 2. Atualiza o self.live_data
                        row_list = list(self.live_data[row])
                        row_list[col] = val
                        self.live_data[row] = row_list
                        
                        # 3. Atualiza o Dashboard (KPI de Ano Máx e Gráfico de Anos)
                        self.update_dashboard_view()
                        return 

                # Fluxo normal: abre o popup de detalhes
                data = {
                    "arquivo": self.table.item(row, 0).text() if self.table.item(row, 0) else "N/A",
                    "autores": self.table.item(row, 1).text() if self.table.item(row, 1) else "N/A",
                    "titulo": self.table.item(row, 2).text() if self.table.item(row, 2) else "N/A",
                    "ano": self.table.item(row, 3).text() if self.table.item(row, 3) else "N/A",
                    "doi": self.table.item(row, 4).text() if self.table.item(row, 4) else "N/A",
                    "jornal": self.table.item(row, 5).text() if self.table.item(row, 5) else "N/A",
                    "abstract": self.table.item(row, 8).text() if self.table.item(row, 8) else "N/A"
                }
                self.show_abstract_popup(data)

            except Exception as e:
                print(f"Erro na interação: {e}")


# =========================================================================
    #  1. MÉTODOS DE FORMATAÇÃO (MOVIDOS PARA O ESCOPO DA CLASSE)
    # =========================================================================
    
    def _inverter_nome(self, nome_completo):
        """Função auxiliar para inverter 'Nome Sobrenome' para 'SOBRENOME, N.'"""
        partes = nome_completo.split(' ')
        if len(partes) <= 1:
            return nome_completo
        
        sobrenome = partes[-1]
        iniciais = ' '.join([p[0] + '.' for p in partes[:-1]])
        return f"{sobrenome}, {iniciais}"

    def format_apa(self, data):
        """Formata no estilo APA 7th ed. com autores invertidos (Sobrenome, N.)"""
        autores_text = data.get('autores', '')
        if not autores_text:
            return ""

        # Divide os autores e aplica a inversão APA (Sobrenome, Inicial.)
        lista_autores = [a.strip() for a in autores_text.split(';')]
        autores_formatados = []
        for autor in lista_autores:
            # Assume que a entrada é "Nome Sobrenome" e converte para "Sobrenome, N."
            autores_formatados.append(self._inverter_nome(autor))
        
        # Junta com vírgulas e usa "&" antes do último
        if len(autores_formatados) > 1:
            autores_final = ", ".join(autores_formatados[:-1]) + " & " + autores_formatados[-1]
        else:
            autores_final = autores_formatados[0]
            
        doi = data.get('doi', '').replace('https://doi.org/', '').replace('http://doi.org/', '')
        titulo = data.get('titulo', '')
        jornal = data.get('jornal', '')
        ano = data.get('ano', '')
        
        # Estrutura APA: Sobrenome, N. (Ano). Título. Jornal. https://doi.org/doi
        return f"{autores_final} ({ano}). {titulo}. {jornal}. https://doi.org/{doi}"

    def format_abnt(self, data):
        """Formata no estilo ABNT listando todos os autores (SOBRENOME, Nome)"""
        autores_text = data.get('autores', '')
        if not autores_text:
            return f"{data.get('titulo', '')}. {data.get('ano', '')}."

        autores_lista = [a.strip() for a in autores_text.split(';')]
        autores_formatados = []

        for autor in autores_lista:
            # Assume entrada "Nome Sobrenome" e converte para "SOBRENOME, Nome"
            partes = autor.split(' ')
            sobrenome = partes[-1].upper()
            nome = ' '.join(partes[:-1])
            autores_formatados.append(f"{sobrenome}, {nome}")

        # Une todos os autores com ponto e vírgula
        autores_final = "; ".join(autores_formatados)

        doi = data.get('doi', '').replace('https://doi.org/', '')
        titulo = data.get('titulo', '')
        jornal = data.get('jornal', '')
        ano = data.get('ano', '')

        return f"{autores_final}. {titulo}. {jornal}, {ano}. DOI: {doi}."



    # =========================================================================
    #  2. POPUP (ATUALIZADO PARA USAR OS MÉTODOS DA CLASSE)
    # =========================================================================

    def show_abstract_popup(self, data):
        dialog = QDialog(self)
        titulo_safe = data['titulo'][:50] if data['titulo'] else "Sem Título"
        dialog.setWindowTitle(f"Detalhes: {titulo_safe}...")
        dialog.setMinimumSize(900, 650)
        dialog.setStyleSheet("background-color: #050505;")
        
        layout = QVBoxLayout(dialog)

        # --- HEADER ---
        header_browser = QTextBrowser()
        header_browser.setOpenExternalLinks(True)
        header_browser.setFrameStyle(QFrame.NoFrame)
        header_browser.setFixedHeight(200) 
        
        html_header = f"""
        <div style='font-family: "Segoe UI", sans-serif; line-height: 1.6; padding: 10px;'>
            <div style='background-color: #111; padding: 10px; border-radius: 5px; border-left: 5px solid #00D2FF; margin-bottom: 5px;'>
                <div style='color: #00D2FF; font-size: 11px; text-transform: uppercase; font-weight: bold; letter-spacing: 1px;'>
        Reference Overview
                </div>
        <h2 style='color: #FFF; margin: 5px 0 10px 0; padding: 0; line-height: 1.2;'>{data['titulo']}</h2>
            <div style='color: #DDD; font-size: 14px;'>{data['autores']}</div>
                <div style='color: #AAA; font-size: 13px; margin-top: 5px;'>
        <i>{data['jornal']}</i> — {data['ano']} | 
        <b>DOI:</b> <a href='{data['doi']}' style='color: #00D2FF; text-decoration: none;'>{data['doi']}</a>
                </div>
            </div>
                <div style='color: #00D2FF; font-size: 11px; margin-top: 15px; text-transform: uppercase; font-weight: bold; letter-spacing: 1px;'>
        Abstract
            </div>
        </div>
        """
        header_browser.setHtml(html_header)
        layout.addWidget(header_browser)

        # --- CAMPO DE EDIÇÃO ---
        abstract_edit = QTextEdit()
        abstract_text = data['abstract'] if data['abstract'] != 'N/A' else ""
        abstract_edit.setPlainText(abstract_text)
        abstract_edit.setPlaceholderText("Insira abstract here...")
        abstract_edit.setStyleSheet("""
            QTextEdit { background-color: #0A0A0A;    color: #DDD;    border: 1px solid #222;    border-radius: 5px;     padding: 15px;        font-family: 'Segoe UI';   font-size: 14px;line-height: 1.8; }
            QTextEdit:focus { border: 1px solid #00D2FF; }
        """)
        layout.addWidget(abstract_edit)

        # --- BOTÕES SUPERIORES ---
        btn_layout = QHBoxLayout()
        btn_save = QPushButton(" Save to Database")
        btn_save.setCursor(Qt.PointingHandCursor)
        btn_save.setStyleSheet("""
            QPushButton { background: #004433; color: #00FFCC; padding: 10px;  border-radius: 5px; border: 1px solid #006644; font-weight: bold;  }
            QPushButton:hover { background: #005544; }
        """)
        
        btn_copy_bib = QPushButton(" Copy BibTeX")
        btn_copy_bib.setCursor(Qt.PointingHandCursor)
        btn_copy_bib.setStyleSheet("""
            QPushButton {  background: #003344; color: #00D2FF; padding: 10px;  border-radius: 5px; border: 1px solid #005577; font-weight: bold;}
            QPushButton:hover { background: #004455; }
        """)

        btn_layout.addWidget(btn_save)
        btn_layout.addWidget(btn_copy_bib)

        # --- FUNÇÕES DE CÓPIA INTERNAS DO POPUP ---
        def copy_apa():
            # CORREÇÃO: Agora chama self.format_apa
            texto = self.format_apa(data)
            QApplication.clipboard().setText(texto)

        def copy_abnt():
            # CORREÇÃO: Agora chama self.format_abnt
            texto = self.format_abnt(data)
            QApplication.clipboard().setText(texto)


        # --- BOTÕES INFERIORES ---
        btn_copy_apa = QPushButton(" Copy APA")
        btn_copy_apa.setCursor(Qt.PointingHandCursor)
        btn_copy_apa.setStyleSheet("""
            QPushButton { background: #222244; color: #99CCFF; padding: 10px; border-radius: 5px; border: 1px solid #444488; font-weight: bold;}
            QPushButton:hover { background: #333355; }
        """)

        btn_copy_abnt = QPushButton(" Copy ABNT")
        btn_copy_abnt.setCursor(Qt.PointingHandCursor)
        btn_copy_abnt.setStyleSheet("""
            QPushButton { background: #442222; color: #FFCC99; padding: 10px; border-radius: 5px; border: 1px solid #884444; font-weight: bold;}
            QPushButton:hover { background: #553333; }
        """)

        btn_copy_apa.clicked.connect(copy_apa)
        btn_copy_abnt.clicked.connect(copy_abnt)

        btn_layout.addWidget(btn_copy_apa)
        btn_layout.addWidget(btn_copy_abnt)
        layout.addLayout(btn_layout)

        # --- LÓGICA DE SALVAR E BIBTEX ---
        def save_changes():
            new_abstract = abstract_edit.toPlainText().strip()
            row_idx = self.table.currentRow()
            
            if row_idx >= 0:
                self.table.setItem(row_idx, 8, QTableWidgetItem(new_abstract))
                
                # Atualiza a memória central
                self.live_data[row_idx][8] = new_abstract if new_abstract else "N/A"
                # Atualiza o dicionário local
                data['abstract'] = new_abstract
            btn_save.setText("✓ Saved!")
            btn_save.setEnabled(False)

        def copy_and_close():
            current_abs = abstract_edit.toPlainText().strip()
            # Tratamento simples para chave de autor
            try:
                author_key = data['autores'].split(',')[0].strip().split(' ')[0]
            except:
                author_key = "Unknown"
            
            bib_key = f"{author_key}{data['ano']}"
            
            bibtex = f"""@article{{{bib_key},
        author = {{{data['autores'].replace(';', ' and')}}},
        title = {{{data['titulo']}}},
        journal = {{{data['jornal']}}},
        year = {{{data['ano']}}},
        doi = {{{data['doi'].replace('https://doi.org/', '')}}},
        abstract = {{{current_abs}}}
        }}"""
            QApplication.clipboard().setText(bibtex)
            dialog.accept()

        btn_save.clicked.connect(save_changes)
        btn_copy_bib.clicked.connect(copy_and_close)

        dialog.exec_()

    # =========================================================================
    #  3. FUNÇÕES DE EXPORTAÇÃO (AGORA ENCONTRAM OS MÉTODOS ACIMA)
    # =========================================================================

    def export_apa(self):
        if not self.live_data:
            return

        linhas = []
        for item in self.live_data:
            data = {
                "autores": item[1],
                "titulo": item[2],
                "jornal": item[3],
                "ano": item[4],
                "doi": item[5],
                "abstract": item[8]
            }
            # CORREÇÃO: Chama self.format_apa corretamente
            linhas.append(self.format_apa(data))

        caminho, _ = QFileDialog.getSaveFileName(self, "Salvar APA", "referencias_APA.txt", "Text (*.txt)")
        if caminho:
            with open(caminho, "w", encoding="utf-8") as f:
                f.write("\n\n".join(linhas))

    def export_abnt(self):
        if not self.live_data:
            return

        linhas = []
        for item in self.live_data:
            data = {
                "autores": item[1],
                "titulo": item[2],
                "jornal": item[3],
                "ano": item[4],
                "doi": item[5],
                "abstract": item[8]
            }
            # CORREÇÃO: Chama self.format_abnt corretamente
            linhas.append(self.format_abnt(data))

        caminho, _ = QFileDialog.getSaveFileName(self, "Salvar ABNT", "referencias_ABNT.txt", "Text (*.txt)")
        if caminho:
            with open(caminho, "w", encoding="utf-8") as f:
                f.write("\n\n".join(linhas))

    def open_pdf(self):
        filters = "Arquivos Suportados (*.pdf *.bib);;PDF (*.pdf);;BibTeX (*.bib)"
        p, _ = QFileDialog.getOpenFileName(self, "Adicionar ao Banco de Dados", "", filters)
        
        if not p:
            return
        ext = os.path.splitext(p)[1].lower()

        if ext == '.pdf':
            local_url = QUrl.fromLocalFile(p)
            self.pdf_viewer.setUrl(local_url)
            self.pdf_viewer.show() 
            self.tabs.setCurrentIndex(0)
            
            self.worker = DeepExtractionWorker(p)
            self.worker.initial_count.connect(self.handle_initial_count)
            self.worker.progress_update.connect(self.on_progress_update)
            self.worker.item_completed.connect(self.append_data_item)
            self.worker.log_result.connect(self.append_log)
            self.worker.finished_signal.connect(self.on_process_finished)
            self.worker.start()

        elif ext == '.bib':
            self.process_bibtex_file(p)

    def handle_initial_count(self, total):
        self.is_processing = True
        self.current_progress = (0, total)
        self.spinner_timer.start(80)

    def on_progress_update(self, current, total):
        self.current_progress = (current, total)
        self.update_progress_label()


    def update_progress_label(self):
            if not self.is_processing:
                self.progress_label.hide()
                return
                
            if self.progress_label.isHidden():
                self.progress_label.show()

            spinner_chars = ["⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏"]
            spinner = spinner_chars[self.spin_idx % len(spinner_chars)]
            
            current, total = self.current_progress
            bar_length = 135 
            
            if total == 0:
                scan_pos = self.spin_idx % bar_length
                bar_list = list(" " * bar_length)
                glow_chars = ["━", "─", "┄", "·"]
                for i, char in enumerate(glow_chars):
                    idx = (scan_pos - i) % bar_length
                    bar_list[idx] = char
                bar = "".join(bar_list)
                msg = f"{spinner}  INITIALIZING CORE  {bar}  {spinner}"
            else:
                pct = current / total
                filled_exact = bar_length * pct
                filled_full = int(filled_exact)
                
                smooth_tip = [" ", "╸", "━"] 
                remainder = int((filled_exact - filled_full) * (len(smooth_tip) - 1))
                tip = smooth_tip[remainder] if filled_full < bar_length else ""
                
                bar_main = "━" * filled_full
                bar_empty = " " * (bar_length - filled_full - (1 if tip else 0))
                
                if current >= total:
                    full_bar = "━" * bar_length
                    msg = f"✔  SYSTEM READY  {full_bar}"
                else:
                    # 1. Descobre quantos dígitos o total tem (ex: 50 -> 2, 1500 -> 4)
                    padding = len(str(total))
                    # 2. Formata dinamicamente: o '0' indica preenchimento com zeros, 
                    current_fmt = format(current, f"0{padding}d")
                    total_fmt = format(total, f"0{padding}d")

                    msg = (
                        f"{current_fmt}/{total_fmt}  "
                        f"{spinner}  ⟪ {bar_main}{tip}{bar_empty} ⟫  {spinner}  "
                        f"{int(pct*100):3}%"
                    )
            
            self.progress_label.setText(msg)
            self.spin_idx += 1

    def append_log(self, html_content):
        self.tab_scan.append(html_content)
        sb = self.tab_scan.verticalScrollBar(); sb.setValue(sb.maximum())

    def append_data_item(self, item_data):
        full_row = [self.worker.file_identifier] + item_data
        self.live_data.append(full_row)
        self.lbl_count.setText(f"{len(self.live_data)} Total refs")
        self.add_row_to_table(full_row)

    def add_row_to_table(self, row_data):
        row = self.table.rowCount()
        self.table.insertRow(row)
        for c, val in enumerate(row_data):
            item = QTableWidgetItem(str(val))
            if c == 4 and str(val).startswith("http"):
                item.setForeground(QColor("#00D2FF")); font = item.font(); font.setUnderline(True); item.setFont(font)
            if c == 0:
                item.setForeground(QColor("#FFCC00"))
            self.table.setItem(row, c, item)

    def on_process_finished(self, success):
        self.is_processing = False
        self.spinner_timer.stop()
        self.progress_label.setText("") 
        msg = "<br><bstyle='padding: 10px; color: #00FF00; font-size: 12px;'>Finish.</b><br>" if success else "<br><b style='color:#FF3366;'>Error.</b>"
        self.tab_scan.append(msg)
        self.update_table_view()

    def update_table_view(self):
        self.table.setRowCount(0)
        self.table.setRowCount(len(self.live_data))
        self.table.setSortingEnabled(False)
        for r, row_data in enumerate(self.live_data):
            for c, val in enumerate(row_data):
                item = QTableWidgetItem(str(val))
                if c == 4 and str(val).startswith("http"):
                    item.setForeground(QColor("#00D2FF")); font = item.font(); font.setUnderline(True); item.setFont(font)
                if c == 0:
                    item.setForeground(QColor("#FFCC00"))
                self.table.setItem(r, c, item)
        self.table.setSortingEnabled(True)
        self.lbl_count.setText(f"{len(self.live_data)} Refs Totais")



    def get_clean_parts(self, name):
            """Prepara o nome e ignora termos genéricos/inválidos."""
            if not name: return []
            
            # Lista de nomes para ignorar completamente
            ignore_list = [
                "unknown", "author", "unknown author", "n/a", "n.a", 
                "none", "anonymous", "et al", "editor", "various"
            ]
            
            # Limpeza inicial
            name_clean = name.lower().replace('.', ' ').strip()
            if name_clean in ignore_list or len(name_clean) < 3:
                return []

            # Inverte se houver vírgula (Sobrenome, Nome -> Nome Sobrenome)
            if ',' in name:
                parts = name.split(',')
                name = f"{parts[1].strip()} {parts[0].strip()}"
            
            name = name.lower().replace('.', ' ')
            parts = [w.strip() for w in name.split() if len(w.strip()) > 0]
            
            # Verifica se o resultado final não é apenas lixo acadêmico
            full_check = " ".join(parts)
            if full_check in ignore_list:
                return []
                
            return parts
    def merge_authors(self, raw_authors):
            """Normaliza nomes de autores (ex: 'Smith, J.' e 'Smith, John' viram o mesmo)."""
            clean_auths = []
            for auth in raw_authors:
                # Remove espaços extras e coloca em Title Case
                a = auth.strip().title()
                if len(a) < 2: continue
                
                # Lógica simples: Pega apenas o sobrenome (antes da vírgula)
                # Para uma análise mais complexa, precisaria de uma lib específica
                # Aqui vamos manter o nome completo limpo para evitar agrupar errados
                clean_auths.append(a)
                
            return clean_auths
    def update_dashboard_view(self):
        if not self.live_data: 
            return

        # Inicializa as listas
        yrs, journals, pubs = [], [], []
        titles_list, raw_titles_only, raw_abstracts_only = [], [], []
        all_raw_authors = [] 
        abs_count = 0
        unique_files = set()

        # Itera sobre os dados carregados
        for row in self.live_data:
            unique_files.add(row[0]) # Coluna 0: File path
            
            # Título (Col 2)
            if row[2] and row[2] != "Sem Título": 
                titles_list.append(row[2])
                raw_titles_only.append(row[2]) 
            
            # Ano (Col 3)
            if str(row[3]).isdigit(): 
                yrs.append(str(row[3]))
            
            # Autores (Col 1)
            if row[1] and row[1] != "N/A":
                # Divide por ponto e vírgula
                all_raw_authors.extend([a.strip() for a in row[1].split(';') if len(a) > 2])
                
            # Journal (Col 5)
            if row[5] and row[5] not in ["", "N/A"]: 
                journals.append(row[5])
            
            # Publisher (Col 6)
            if row[6] and row[6] not in ["", "N/A"]: 
                pubs.append(row[6])
                
            # Abstract (Col 8)
            if row[8] and row[8] != "N/A":
                abs_count += 1
                raw_abstracts_only.append(row[8]) 

        # --- Atualização dos KPIs (Texto grande) ---
        self.k1.v.setText(str(len(self.live_data)))
        self.k2.v.setText(str(abs_count))
        self.k4.v.setText(str(len(unique_files)))
        if yrs: 
            self.k3.v.setText(str(max(yrs)))
        else:
            self.k3.v.setText("-")

        # --- Atualização dos Gráficos Circulares e Barras ---
        # Certifique-se que seus objetos de gráfico (DonutChart, BarChart) têm o método .update_data()
        self.c_yrs.update_data(Counter(yrs))
        # Coluna 7 é TYPE
        self.c_tps.update_data(Counter([row[7] for row in self.live_data]))
        self.c_pub.update_data(Counter(pubs))
        self.b_journal.update_data(Counter(journals))
        self.b_top_articles.update_data(Counter(titles_list))
        
        # Autores unificados
        merged_auths_dict = self.merge_authors(all_raw_authors)
        self.b_aut.update_data(Counter(merged_auths_dict))
        
        # --- LÓGICA DE KEYWORDS DUPLA ---
        # 1. Tenta pegar o que está nas tabelas da Aba 4 (marcados pelo usuário)
        # Nota: get_active_keywords agora retorna dois valores (veja passo 1 acima)
        title_kw, abs_kw = self.get_active_keywords()
        
        # 2. Fallback: se as tabelas estiverem vazias (primeira execução), extrai do zero
        if not title_kw:
            title_kw = self.get_top_terms(raw_titles_only, min_freq=2, min_len=4)
        if not abs_kw:
            abs_kw = self.get_top_terms(raw_abstracts_only, min_freq=2, min_len=5)
            
        # 3. Plota nos dois gráficos de barra separados no Dashboard
        # b_word_titles e b_word_freq devem ser instâncias de BarChart criadas no setup_dash_widgets
        self.b_word_titles.update_data(title_kw)
        self.b_word_freq.update_data(abs_kw)


    def normalize_author_name(self, name):
            """Limpa o nome, remove pontuações e inverte 'Sobrenome, Nome' para 'Nome Sobrenome'."""
            name = str(name).strip()
            invalid_terms = ["unknown", "author", "n/a", "none", "et al", "others"]
            if any(term in name.lower() for term in invalid_terms) or len(name) < 2:
                return None
            
            # Se estiver no formato "Sobrenome, Nome", inverte para "Nome Sobrenome"
            if ',' in name:
                parts = name.split(',', 1)
                name = f"{parts[1].strip()} {parts[0].strip()}"
                
            # Remove pontos, aspas e vírgulas restantes (ex: "J. Smith" -> "J Smith")
            name = re.sub(r'[.,"\']', '', name)
            
            # Pega as palavras, capitaliza as primeiras letras e remove espaços extras
            words = [w.title() for w in name.split() if w]
            if not words:
                return None
                
            return " ".join(words)

    def get_active_keywords(self):
        """Lê as checkboxes das DUAS tabelas e retorna dois Counters separados."""
        
        def extract_from(table):
            data = {}
            for r in range(table.rowCount()):
                # Verifica se a checkbox na coluna 0 está marcada
                item = table.item(r, 0)
                if item and item.checkState() == Qt.Checked:
                    # Palavra na coluna 1, Frequência na coluna 2
                    word = table.item(r, 1).text()
                    try:
                        freq = int(table.item(r, 2).text())
                        data[word] = freq
                    except:
                        continue
            return Counter(data)

        # Retorna uma tupla: (Counter Títulos, Counter Abstracts)
        return extract_from(self.kw_table_titles), extract_from(self.kw_table_abstracts)
    def export_bibtex(self):
        path, _ = QFileDialog.getSaveFileName(self, "Exportar BibTeX Consolidado", "", "*.bib")
        if path and self.live_data:
            try:
                with open(path, "w", encoding="utf-8") as f:
                    for r in self.live_data:
                        # r[1] = Autor, r[3] = Ano, r[8] = Abstract
                        # Gerar uma chave de citação simples
                        last_name = re.sub(r'\W+', '', r[1].split(',')[0].split(' ')[0])
                        k = f"{last_name}{r[3]}"
                        
                        # Limpeza do abstract para não quebrar o BibTeX
                        raw_abs = str(r[8])
                        clean_abs = raw_abs.replace('\n', ' ').replace('\r', '').strip()
                        abs_field = f"  abstract={{{clean_abs}}},\n" if clean_abs not in ["N/A", ""] else ""
                        
                        f.write(f"@article{{{k},\n"
                                f"  author={{{r[1]}}},\n"
                                f"  title={{{r[2]}}},\n"
                                f"  journal={{{r[5]}}},\n"
                                f"  year={{{r[3]}}},\n"
                                f"  doi={{{r[4]}}},\n"
                                f"{abs_field}"
                                f"  publisher={{{r[6]}}},\n"
                                f"  note={{Source File: {r[0]}}}\n"
                                f"}}\n\n")
                QMessageBox.information(self, "Exportar", "Arquivo BibTeX salvo com sucesso!")
            except Exception as e:
                QMessageBox.critical(self, "Erro", f"Falha ao salvar arquivo: {e}")

    def export_excel(self):
        path, _ = QFileDialog.getSaveFileName(self, "Exportar Excel Consolidado", "", "*.xlsx")
        if path and self.live_data:
            df = pd.DataFrame(self.live_data, columns=["Arquivo Fonte", "Autores", "Título", "Ano", "DOI", "Jornal", "Publisher", "Tipo", "Abstract"])
            df.to_excel(path, index=False)
            QMessageBox.information(self, "Exportar", "Arquivo Excel salvo.")
                
    def process_bibtex_file(self, fname):
            """Inicia o Worker de BibTeX para processamento em background."""
            if getattr(self, 'is_processing', False):
                return  # Evita iniciar dois processos ao mesmo tempo

            self.is_processing = True
            
            # OTIMIZAÇÃO: Desativa a ordenação UMA VEZ antes de começar
            self.table.setSortingEnabled(False) 

            self.bib_worker = BibtexWorker(fname)
            
            # Conexões
            self.bib_worker.initial_count.connect(self.handle_initial_count)
            self.bib_worker.header_ready.connect(self.append_log) # Assumindo que append_log aceita HTML
            self.bib_worker.log_result.connect(self.append_log)
            self.bib_worker.progress_update.connect(self.handle_bib_progress)
            self.bib_worker.item_completed.connect(self.add_single_item_from_worker)
            self.bib_worker.finished_signal.connect(self.on_bib_finished)
            
            self.bib_worker.start()

    def handle_bib_progress(self, current, total):
            self.current_progress = (current, total)
            self.update_progress_label() # Assumindo que este método atualiza a barra/texto

    def add_single_item_from_worker(self, full_row):
            """
            Atualiza o banco de dados e a tabela automaticamente.
            Nota: Não ativamos o sort aqui para ganhar velocidade na inserção.
            """
            self.live_data.append(full_row)
            self.lbl_count.setText(f"{len(self.live_data)} Total refs")
            
            # Apenas adiciona a linha. A ordenação está desligada globalmente durante o processo.
            self.add_row_to_table(full_row)

    def on_bib_finished(self, total):
            """Finaliza o processo de importação BibTeX."""
            self.is_processing = False
            
            # OTIMIZAÇÃO: Reativa a ordenação apenas no final de tudo
            self.table.setSortingEnabled(True)

            if total > 0:
                self.append_log(f"<div style='color:#00FF00; padding:10px; font-weight:bold;'>"
                                f"✓ Importação Automática Concluída: {total} itens adicionados.</div>")
            elif total == 0:
                self.append_log("<div style='color:#FFCC00; padding:10px;'>Aviso: Nenhum item válido encontrado no arquivo .bib.</div>")
            
            # Feedback visual de erro se total for -1 (sinal de erro no worker)
            if total == -1:
                self.append_log("<div style='color:#FF3366; padding:10px;'>Erro crítico ao ler o arquivo.</div>")

            QTimer.singleShot(2000, self.stop_processing_ui)

    def stop_processing_ui(self):
            """Reseta os flags de animação."""
            self.is_processing = False
            self.current_progress = (0, 0)
            self.update_progress_label()

    def run_author_extraction(self):
        if not self.live_data: return
        
        all_authors = []
        for row in self.live_data:
            auth_str = str(row[1]) if row[1] else ""
            # Divide por ; ou , (comuns em BibTeX e PDFs)
            parts = re.split(r'[;]', auth_str) if ';' in auth_str else re.split(r',', auth_str)
            
            for p in parts:
                norm = self.normalize_author_name(p)
                if norm:
                    all_authors.append(norm)

        counts = Counter(all_authors)
        # Preenche a tabela auth_table (que você criou no passo anterior)
        self.fill_kw_table(self.auth_table, counts)

    def normalize_author_name(self, name):
        # 1. Limpeza básica
        name = name.strip().upper()
        # 2. Filtro de "Unknown"
        invalid_terms = ["UNKNOWN", "AUTHOR", "N/A", "NONE", "ET AL", "OTHERS"]
        if any(term in name for term in invalid_terms) or len(name) < 3:
            return None
        
        # 3. Remove caracteres especiais que atrapalham a unificação
        name = name.replace(".", "").replace(",", ", ") # Garante espaço após vírgula
        name = " ".join(name.split()) # Remove espaços duplos
        return name
    

    def run_author_network_analysis(self):
            if not self.live_data:
                QMessageBox.warning(self, "Aviso", "Sem dados carregados.")
                return
            
            dialog = GraphConfigDialog(self)
            if dialog.exec_() == QDialog.Accepted:
                config = dialog.get_values()
                
                papers_as_author_lists = []
                
                for row in self.live_data:
                    auth_str = str(row[1]) if row[1] else ""
                    
                    # 1. Padroniza a separação: troca ' and ' por ';'
                    auth_str = re.sub(r'\s+and\s+', ';', auth_str, flags=re.IGNORECASE)
                    
                    # 2. Divide os autores. Damos preferência ao ';' para não destruir nomes formatados com vírgula.
                    if ';' in auth_str:
                        raw_authors = auth_str.split(';')
                    else:
                        raw_authors = auth_str.split(',')
                    
                    valid_names = []
                    for p in raw_authors:
                        norm = self.normalize_author_name(p)
                        if norm:
                            # 3. O SEGREDO: Troca o espaço por '_' (ex: "Joao_Silva").
                            # Isso impede o CountVectorizer de quebrar o nome da pessoa na metade!
                            valid_names.append(norm.replace(" ", "_"))
                    
                    if valid_names:
                        # Junta todos os autores que participaram deste paper em uma única string
                        papers_as_author_lists.append(" ".join(valid_names))

                # Gera o grafo
                self.auth_network_widget.generate_graph(
                    papers_as_author_lists, 
                    min_freq=1, 
                    top_n=config.get("top_n", 50), 
                    threshold=config["threshold"],
                    font_color=config["font_color"],
                    layout_type=config["layout"]
                )

if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = BiblioApp()
    w.show()
    sys.exit(app.exec_())