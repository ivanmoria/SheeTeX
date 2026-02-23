import sys, re, webbrowser, math, html
import pandas as pd
import os
import platform
import subprocess
from collections import Counter

# PyQt5 Imports
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QPushButton, QTextBrowser, QLabel, QFileDialog, 
                             QTabWidget, QTableWidget, QTableWidgetItem, 
                             QHeaderView, QHBoxLayout, QFrame, QScrollArea, QToolTip,
                             QSplitter)
from PyQt5.QtCore import QThread, pyqtSignal, Qt, QRectF, QTimer, QPoint, QUrl
from PyQt5.QtGui import QPainter, QColor, QFont, QTextCursor, QDesktopServices
from PyQt5.QtWebEngineWidgets import QWebEngineView, QWebEngineSettings 

# PDF & Data Imports
from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextContainer, LAParams
from habanero import Crossref

# =============================================================================
# ENGINE: PROCESSAMENTO DE PDF (WORKER)
# Integração: Lógica de leitura robusta + Logs ricos e layout do Código 2
# =============================================================================
class DeepExtractionWorker(QThread):
    initial_count = pyqtSignal(int)
    status_update = pyqtSignal(str) 
    item_completed = pyqtSignal(list)
    log_result = pyqtSignal(str)   
    finished_signal = pyqtSignal(bool)

    def __init__(self, pdf_path):
        super().__init__()
        self.pdf_path = pdf_path

    def clean_text(self, text):
        if not text: return ""
        text = html.unescape(text)
        return re.sub(r'\s+', ' ', text).strip().strip(',').strip('.')

    def clean_abstract(self, text):
        """Remove tags XML/HTML comuns em abstracts do Crossref"""
        if not text: return "N/A"
        text = html.unescape(text)
        clean = re.sub(r'<[^>]+>', '', text)
        return re.sub(r'\s+', ' ', clean).strip()

    def format_person(self, p):
        fn = p.get('given', '').strip()
        ln = p.get('family', p.get('name', '')).strip() 
        return f"{self.clean_text(ln)}, {self.clean_text(fn)}" if ln and fn else self.clean_text(ln)
    
    def run(self):
        cr = Crossref()
        try:
            self.status_update.emit("Lendo estrutura do arquivo PDF...")
            laparams = LAParams(char_margin=3.5)
            full_text = ""
            first_page_text = ""
            
            # --- 1. LEITURA DO PDF (Estabilidade do Cod 1) ---
            page_count = 0
            for page in extract_pages(self.pdf_path, laparams=laparams):
                page_text = ""
                for el in page:
                    if isinstance(el, LTTextContainer): 
                        page_text += el.get_text()
                
                if page_count == 0:
                    first_page_text = page_text
                
                full_text += page_text
                page_count += 1

            # --- 2. IDENTIFICAÇÃO DOS DADOS DO ARQUIVO (Layout Rico do Cod 2) ---
            self.status_update.emit("Identificando metadados do documento base...")
            
            doc_title = os.path.basename(self.pdf_path)
            doc_author = "Autor Desconhecido"
            doc_year = "n.d."
            doc_journal = "Periódico Não Identificado"
            doc_doi_link = ""
            doc_full_ref = "Referência completa não disponível."

            # Tenta achar DOI na primeira página
            main_doi_match = re.search(r"10\.\d{4,9}/[-._;()/:A-Z0-9]+", first_page_text, re.IGNORECASE)
            main_item = None
            if main_doi_match:
                try: main_item = cr.works(ids=main_doi_match.group(0).rstrip('.'))['message']
                except: pass
            
            # Fallback: Busca por texto
            if not main_item:
                q_clean = re.sub(r'\s+', ' ', first_page_text[:300]).strip()
                try:
                    res = cr.works(query=q_clean, limit=1)
                    if res['message']['items']: main_item = res['message']['items'][0]
                except: pass

            if main_item:
                doc_title = main_item.get("title", [doc_title])[0]
                auths = main_item.get('author', [])
                if auths:
                    doc_author = "; ".join([f"{a.get('family', '')}, {a.get('given', '')}" for a in auths])
                
                date_parts = main_item.get("published-print", {}).get("date-parts") or \
                             main_item.get("published-online", {}).get("date-parts") or \
                             main_item.get("created", {}).get("date-parts")
                if date_parts: doc_year = str(date_parts[0][0])
                
                doc_journal = main_item.get("container-title", [doc_journal])[0]
                doi_val = main_item.get("DOI", "")
                doc_doi_link = f"https://doi.org/{doi_val}" if doi_val else ""
                
                doc_full_ref = f"<i>{doc_journal}</i>  • {doc_year}. "
                if doc_doi_link:
                    doc_full_ref += f"  •  DOI: <a href='{doc_doi_link}' style='color:#00D2FF; text-decoration:none;'>{doi_val}</a>"

            # --- 3. LOG HEADER (Visual do Cod 2) ---
            header_html = (
                f"<div style='padding: 15px 0; border-bottom: 2px solid #333; margin-bottom: 20px; font-family: sans-serif; line-height: 1.5;'>"
                f"  <div style='margin-bottom: 8px; display:flex; justify-content:space-between;'>"
                f"      <span style='color: #00D2FF; font-weight: bold; font-size: 12px; text-transform: uppercase; letter-spacing: 1px;'>Documento Base</span>"
                f"  </div>"
                f"  <div style='color: #FFF; font-size: 14px; font-weight: bold; margin-bottom: 6px;'>{doc_title}</div>"
                f"  <div style='color: #CCC; font-size: 12px;'>{doc_author}</div>"
                f"  <div style='margin-top: 10px; padding-top: 10px; border-top: 1px dotted #333; font-size: 12px; color: #999;'>"
                f"      {doc_full_ref}"
                f"  </div>"
                f"</div>"
            )
            self.log_result.emit(header_html)

            # --- 4. EXTRAÇÃO DE REFERÊNCIAS ---
            content = full_text
            for kw in ["references", "bibliography", "referências", "bibliografia"]:
                found = full_text.lower().rfind(kw)
                if found != -1: content = full_text[found:]; break
            
            # Regex robusto para separar referências
            segments = [s.strip().replace('\n', ' ') for s in re.split(r"\n(?=\[?\d+\]?[\.\s]|\b[A-Z][a-z]+, [A-Z]\.)", content) if len(s.strip()) > 15]
            
            total = len(segments)
            self.initial_count.emit(total)
            history = {} 
            base_item_style = "padding: 12px 0; border-bottom: 1px solid #222; font-family: sans-serif; font-size: 12px; line-height: 1.4;"

            for i, seg in enumerate(segments):
                idx = i + 1
                self.status_update.emit(f"Processando referência {idx}/{total}...")
                
                try:
                    doi_m = re.search(r"10\.\d{4,9}/[-._;()/:A-Z0-9]+", seg, re.IGNORECASE)
                    item = None
                    if doi_m:
                        try: item = cr.works(ids=doi_m.group(0).rstrip('.'))['message']
                        except: item = None
                    
                    if not item:
                        q = re.sub(r'[^\w\s]', '', seg[:160])
                        res = cr.works(query=q, limit=1)
                        if res['message']['items']: item = res['message']['items'][0]

                    if item:
                        auths = item.get('author', item.get('editor', []))
                        auth_str = "; ".join([self.format_person(a) for a in auths]) or "Unknown Author"
                        year = str(item.get("created", {}).get("date-parts", [[None]])[0][0])
                        title = self.clean_text(item.get("title", ["Unknown Title"])[0])
                        journal = self.clean_text(item.get("container-title", ["N/A"])[0])
                        res_doi = item.get('DOI', '').lower()
                        link = f"https://doi.org/{res_doi}" if res_doi else ""
                        uid = res_doi if res_doi else f"{auth_str}{title}{year}".lower()
                        
                        abstract_raw = item.get('abstract', '')
                        abstract_clean = self.clean_abstract(abstract_raw)

                        is_duplicate = uid in history
                        
                        # Definição de Cores para o Log (Cod 2)
                        if is_duplicate:
                            idx_color = "#FF3366" 
                            msg_dup = "<span style='color:#FF3366; font-weight:bold;'> [Duplicada]</span>"
                            main_color = "#777"
                            sub_color = "#555"
                            doi_style = "color:#555; text-decoration:none; cursor: default;"
                        else:
                            idx_color = "#00D2FF" 
                            msg_dup = ""
                            main_color = "#FFF"
                            sub_color = "#BBB"
                            doi_style = "color:#00D2FF; text-decoration:underline;"

                        doi_html = f"<a href='{link}' style='{doi_style}'>{res_doi}</a>" if link else "N/A"
                        
                        # HTML Rico do Log
                        detail_log = (
                            f"<div style='{base_item_style}'>"
                            f"  <div style='margin-bottom: 4px;'>"
                            f"    <span style='color:{idx_color}; font-family:Consolas; font-weight:bold;'>[{idx:03d}/{total:03d}]</span>{msg_dup} "
                            f"    <b style='color:{main_color}; font-size: 13px;'>{title}</b>"
                            f"  </div>"
                            f"  <div style='color:{sub_color};'>{auth_str}</div>"
                            f"  <div style='color:#777; margin-top: 2px;'>{journal} • {year} • DOI: {doi_html}</div>"
                            f"</div>"
                        )

                        if not is_duplicate:
                            history[uid] = title
                            pub = self.clean_text(item.get("publisher", "N/A"))
                            item_type = item.get("type", "other").title() if item.get("type") else "Other"
                            self.item_completed.emit([auth_str, title, year, link, journal, pub, item_type, abstract_clean])
                        
                        self.log_result.emit(detail_log)
                    else:
                        error_log = (
                            f"<div style='{base_item_style} color:#FF3366;'>"
                            f"  <span style='font-family:Consolas; font-weight:bold;'>[{idx:03d}/{total:03d}]</span> "
                            f"  Referência não identificada automaticamente."
                            f"</div>"
                        )
                        self.log_result.emit(error_log)
                
                except Exception as e:
                    except_log = (
                        f"<div style='{base_item_style} color:#FF3366;'>"
                        f"  <span style='font-family:Consolas; font-weight:bold;'>[{idx:03d}/{total:03d}]</span> "
                        f"  Erro ao processar: {str(e)}"
                        f"</div>"
                    )
                    self.log_result.emit(except_log)
            
            self.finished_signal.emit(True)
        except Exception as e: 
            print(f"Erro Fatal: {e}")
            self.finished_signal.emit(False)

# =============================================================================
# CLASSES DE GRÁFICOS (Interativos do Código 2)
# =============================================================================
class DonutChart(QWidget):
    def __init__(self, title):
        super().__init__(); self.title, self.data = title, []
        self.setMinimumHeight(320)
        self.setMouseTracking(True)
        self.colors = [QColor("#00D2FF"), QColor("#7B2FF7"), QColor("#00FFAA"), QColor("#FFCC00"), QColor("#FF3366"), QColor("#FFFFFF")]
    
    def update_data(self, d): 
        self.data = sorted(d.items(), key=lambda x: x[1], reverse=True)[:6]
        self.update()
    
    def mouseMoveEvent(self, event):
        if not self.data: return
        pos, center = event.pos(), QPoint(self.width()//2, self.height()//2 + 10)
        dist = math.sqrt((pos.x()-center.x())**2 + (pos.y()-center.y())**2)
        if 60 < dist < 120:
            angle = math.degrees(math.atan2(-(pos.y()-center.y()), pos.x()-center.x()))
            if angle < 0: angle += 360
            total, start = sum(v for k,v in self.data), 90
            for label, value in self.data:
                span = (value/total) * 360
                norm_angle = (360 - angle + 90) % 360
                if start <= norm_angle <= start + span:
                    QToolTip.setFont(QFont("Segoe UI", 11))
                    QToolTip.showText(event.globalPos(), f"<b>{label}</b>: {value}")
                    return
                start += span
        QToolTip.hideText()

    def paintEvent(self, event):
        p = QPainter(self); p.setRenderHint(QPainter.Antialiasing)
        p.setPen(QColor("#00D2FF")); p.setFont(QFont("Segoe UI", 12, QFont.Bold))
        p.drawText(0, 0, self.width(), 30, Qt.AlignCenter, self.title.upper())
        if not self.data: return
        size = 220
        rect = QRectF(self.width()//2 - size//2, self.height()//2 - size//2 + 10, size, size)
        total, start = sum(v for k,v in self.data), 90*16
        for i, (l, v) in enumerate(self.data):
            span = -int((v/total)*360*16)
            p.setBrush(self.colors[i%len(self.colors)]); p.setPen(Qt.NoPen)
            p.drawPie(rect, start, span); start += span
        p.setBrush(QColor("#050505"))
        p.drawEllipse(self.width()//2 - 120//2, self.height()//2 - 120//2 + 10, 120, 120)
        p.setPen(QColor("#FFF")); p.setFont(QFont("Segoe UI", 18, QFont.Bold))
        p.drawText(rect, Qt.AlignCenter, str(total))

class BarChart(QWidget):
    def __init__(self, title, color_hex="#00D2FF"):
        super().__init__(); self.title, self.data, self.rects = title, [], []
        self.bar_color = QColor(color_hex)
        self.setMinimumHeight(300)
        self.setMouseTracking(True)
        self.margin_left = 200

    def update_data(self, d): self.data = d.most_common(8); self.update()
    
    def mouseMoveEvent(self, event):
        for rect, label, value in self.rects:
            if rect.contains(event.pos()):
                QToolTip.setFont(QFont("Segoe UI", 11))
                QToolTip.showText(event.globalPos(), f"<b>{label}</b>: {value}")
                return
        QToolTip.hideText()
        
    def paintEvent(self, event):
        p = QPainter(self); p.setRenderHint(QPainter.Antialiasing)
        p.setPen(QColor("#00D2FF")); p.setFont(QFont("Segoe UI", 12, QFont.Bold))
        p.drawText(self.margin_left, 0, self.width()-self.margin_left, 30, Qt.AlignLeft, self.title.upper())
        if not self.data: return
        max_v = self.data[0][1]; start_y, bar_h, gap = 40, 20, 10; self.rects = []
        for l, v in self.data:
            w = int((v/max_v)*(self.width() - self.margin_left - 50))
            r = QRectF(self.margin_left, start_y, w, bar_h)
            self.rects.append((r, l, v))
            p.setBrush(self.bar_color); p.setPen(Qt.NoPen); p.drawRect(r)
            p.setPen(QColor("#AAA")); p.setFont(QFont("Segoe UI", 9))
            label_text = str(l)
            if len(label_text) > 30: label_text = label_text[:28] + "..."
            p.drawText(5, int(start_y + bar_h/2 + 4), self.margin_left - 10, bar_h, Qt.AlignRight, label_text)
            p.setPen(QColor("#FFF")); p.setFont(QFont("Segoe UI", 9, QFont.Bold))
            p.drawText(int(self.margin_left + w + 10), int(start_y + bar_h - 4), str(v))
            start_y += (bar_h + gap)

# =============================================================================
# INTERFACE PRINCIPAL (Layout do Código 2 com PDF Settings corrigidos)
# =============================================================================
class BiblioApp(QMainWindow):
    def __init__(self):
        super().__init__(); self.live_data = []
        self.is_processing = False
        self.current_status_text = "Inicializando..."
        self.spinner_chars = ["⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏"]
        self.spin_idx = 0
        self.initUI()
        self.spinner_timer = QTimer()
        self.spinner_timer.timeout.connect(self.update_spinner_animation)

    def initUI(self):
        self.setWindowTitle("BiblioTech Ultimate Integrated V2"); self.resize(1600, 1000)
        self.setStyleSheet("""
            QMainWindow, QWidget { background: #050505; color: #FFF; font-family: 'Segoe UI', Arial; }
            QTableWidget { background: #020202; gridline-color: #111; color: #CCC; border: none; font-size: 13px; }
            QHeaderView::section { background: #0A0A0A; color: #00D2FF; font-weight: bold; padding: 12px; border: none; }
            QPushButton { background: transparent; border: none; padding: 8px 20px; font-weight: bold; font-size: 12px; min-width: 120px; }
            QPushButton:hover { background: #111; border-radius: 4px; }
            QTabWidget::pane { border-top: 1px solid #151515; top: -1px; } 
            QTabBar::tab { background: #0A0A0A; padding: 12px 30px; color: #444; font-weight: bold; border-bottom: 2px solid transparent; }
            QTabBar::tab:selected { color: #FFF; border-bottom: 2px solid #00D2FF; }
            QScrollBar:vertical { background: #050505; width: 12px; }
            QScrollBar::handle:vertical { background: #333; min-height: 20px; border-radius: 6px; }
            QSplitter::handle { background: #111; }
        """)

        self.tabs = QTabWidget()
        self.tabs.setDocumentMode(True)
        self.tabs.currentChanged.connect(self.toggle_buttons)
        self.setCentralWidget(self.tabs)

        # Toolbar Esquerda
        self.left_widget = QWidget()
        self.left_layout = QHBoxLayout(self.left_widget)
        self.left_layout.setContentsMargins(10, 5, 0, 5)
        
        self.btn_load = QPushButton("IMPORTAR PDF")
        self.btn_load.setStyleSheet("color: #00D2FF; border: 1px solid #003344; border-radius: 4px;")
        self.btn_load.clicked.connect(self.open_pdf)
        
        self.left_layout.addWidget(self.btn_load)
        self.tabs.setCornerWidget(self.left_widget, Qt.TopLeftCorner)

        # Toolbar Direita
        self.right_widget = QWidget()
        self.right_layout = QHBoxLayout(self.right_widget)
        self.right_layout.setContentsMargins(0, 5, 20, 5) 
        self.right_layout.setSpacing(10)
        self.btn_bib = QPushButton("EXPORT BIBTEX"); self.btn_bib.setStyleSheet("color: #7B2FF7;")
        self.btn_bib.clicked.connect(self.export_bibtex)
        self.btn_xlsx = QPushButton("EXPORT EXCEL"); self.btn_xlsx.setStyleSheet("color: #00FFAA;")
        self.btn_xlsx.clicked.connect(self.export_excel)
        self.right_layout.addStretch() 
        self.right_layout.addWidget(self.btn_bib); self.right_layout.addWidget(self.btn_xlsx)
        self.tabs.setCornerWidget(self.right_widget, Qt.TopRightCorner)


# ABA 1: SCANNER + PDF
        self.tab1_widget = QWidget()
        layout1 = QHBoxLayout(self.tab1_widget)
        layout1.setContentsMargins(0, 0, 0, 0)
        
        self.splitter = QSplitter(Qt.Horizontal)
        # Define a cor da "alça" do divisor para não sumir no preto
        self.splitter.setStyleSheet("QSplitter::handle { background: #111; }")

        # 1. Painel de Log (60%)
        self.tab_scan = QTextBrowser()
        self.tab_scan.setStyleSheet("background-color: #050505; border: none; padding: 25px;")
        
        # 2. Leitor de PDF (40%)
        self.pdf_viewer = QWebEngineView()
        
        # Inicia com fundo preto absoluto
        self.pdf_viewer.page().setBackgroundColor(QColor("#050505"))
        self.pdf_viewer.setStyleSheet("background-color: #050505; border-left: 1px solid #1a1a1a;")
        self.pdf_viewer.setHtml("<html><body style='background-color: #050505;'></body></html>")

        # Configurações do Engine
        s = self.pdf_viewer.settings()
        s.setAttribute(QWebEngineSettings.PluginsEnabled, True)
        s.setAttribute(QWebEngineSettings.PdfViewerEnabled, True)
        s.setAttribute(QWebEngineSettings.LocalContentCanAccessFileUrls, True)
        
        self.pdf_viewer.setMinimumWidth(100) 

        self.splitter.addWidget(self.tab_scan)
        self.splitter.addWidget(self.pdf_viewer)

        # --- PROPORÇÃO 60/40 ---
        # setSizes define os pixels iniciais (calculados sobre uma base de 1000)
        self.splitter.setSizes([600, 400]) 
        # setStretchFactor garante que ao esticar a janela, a proporção 6:4 se mantenha
        self.splitter.setStretchFactor(0, 6) 
        self.splitter.setStretchFactor(1, 4) 

        layout1.addWidget(self.splitter)

        # ABA 2: BANCO
        self.table = QTableWidget(0, 8)
        self.table.setHorizontalHeaderLabels(["AUTORES", "TÍTULO", "ANO", "DOI", "JORNAL", "PUBLISHER", "TIPO", "ABSTRACT"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setStyleSheet("QTableWidget { background: #0A0A0A; gridline-color: #222; }")

        # ABA 3: DASHBOARD
        self.dash_scroll = QScrollArea()
        self.dash_inner = QWidget()
        self.l_dash = QVBoxLayout(self.dash_inner)
        self.setup_dash_widgets()
        self.dash_scroll.setWidget(self.dash_inner); self.dash_scroll.setWidgetResizable(True)

        self.tabs.addTab(self.tab1_widget, "SCANNER LOG")
        self.tabs.addTab(self.table, "BANCO DE DADOS")
        self.tabs.addTab(self.dash_scroll, "DASHBOARD")
    def setup_dash_widgets(self):
        self.l_dash.setContentsMargins(40, 40, 40, 40); self.l_dash.setSpacing(40)
        h_kpi = QHBoxLayout(); h_kpi.setSpacing(20)
        self.k1 = self.create_kpi("TOTAL REFERÊNCIAS")
        self.k2 = self.create_kpi("ABSTRACTS OBTIDOS") 
        self.k3 = self.create_kpi("ANO REC.")
        self.k4 = self.create_kpi("JORNAIS DISTINTOS")
        h_kpi.addWidget(self.k1); h_kpi.addWidget(self.k2); h_kpi.addWidget(self.k4); h_kpi.addWidget(self.k3)
        self.l_dash.addLayout(h_kpi)

        h_donuts = QHBoxLayout(); h_donuts.setSpacing(20)
        self.c_yrs = DonutChart("Por Ano")
        self.c_tps = DonutChart("Tipos de Documento")
        self.c_pub = DonutChart("Top Publishers")
        h_donuts.addWidget(self.c_yrs); h_donuts.addWidget(self.c_tps); h_donuts.addWidget(self.c_pub)
        self.l_dash.addLayout(h_donuts)

        h_bars = QHBoxLayout(); h_bars.setSpacing(20)
        self.b_journal = BarChart("Top Jornais / Conferências", "#FF3366")
        self.b_aut = BarChart("Autores Mais Citados", "#7B2FF7")
        h_bars.addWidget(self.b_journal); h_bars.addWidget(self.b_aut)
        self.l_dash.addLayout(h_bars)
        
        h_words = QHBoxLayout(); h_words.setSpacing(20)
        self.b_words = BarChart("Termos nos Títulos", "#00FFAA")
        self.b_abstract_words = BarChart("Termos nos Abstracts", "#FFCC00")
        h_words.addWidget(self.b_words); h_words.addWidget(self.b_abstract_words)
        self.l_dash.addLayout(h_words)

    def create_kpi(self, t):
        f = QFrame(); f.setStyleSheet("background: #0F0F0F; border: 1px solid #222; border-radius: 12px;")
        f.setFixedHeight(110)
        l = QVBoxLayout(f); 
        title = QLabel(t); title.setStyleSheet("color: #888; font-size: 13px; font-weight: bold;")
        title.setAlignment(Qt.AlignCenter)
        v = QLabel("-"); v.setStyleSheet("color: #FFF; font-size: 32px; font-weight: 900;")
        v.setAlignment(Qt.AlignCenter)
        l.addWidget(title); l.addWidget(v); f.v = v; return f

    def toggle_buttons(self, idx):
        self.btn_bib.setVisible(idx == 1)
        self.btn_xlsx.setVisible(idx == 2)

    def handle_cell_click(self, row, col):
        if col == 3: # Link DOI
            url = self.table.item(row, col).text()
            if url.startswith("http"): webbrowser.open(url)

    def open_pdf(self):
        p, _ = QFileDialog.getOpenFileName(self, "Selecionar PDF", "", "*.pdf")
        if p:
            local_url = QUrl.fromLocalFile(p)
            self.pdf_viewer.setUrl(local_url)
            self.pdf_viewer.show() 
            self.tab_scan.clear(); self.live_data = []
            self.tabs.setCurrentIndex(0)
            
            self.worker = DeepExtractionWorker(p)
            self.worker.initial_count.connect(self.handle_initial_count)
            self.worker.status_update.connect(self.update_status_text)
            self.worker.item_completed.connect(self.live_data.append)
            self.worker.log_result.connect(self.append_final_log)
            self.worker.finished_signal.connect(self.on_process_finished)
            self.worker.start()

    def handle_initial_count(self, total):
        header = (
            f"<div style='margin-bottom: 20px; font-size: 14px; color: #FFF;'>"
            f"<b>SISTEMA INICIADO</b><br>"
            f"<span style='color: #00D2FF;'>Encontradas {total} referências bibliográficas.</span><br>"
            f"</div><br>" 
        )
        cursor = self.tab_scan.textCursor()
        cursor.insertHtml(header)
        cursor.movePosition(QTextCursor.End)
        self.tab_scan.setTextCursor(cursor)
        self.is_processing = True
        self.spinner_timer.start(80)

    def update_status_text(self, text): self.current_status_text = text

    def update_spinner_animation(self):
        if not self.is_processing: return
        cursor = self.tab_scan.textCursor()
        cursor.movePosition(QTextCursor.End)
        cursor.select(QTextCursor.LineUnderCursor)
        cursor.removeSelectedText()
        char = self.spinner_chars[self.spin_idx % len(self.spinner_chars)]
        html_loader = f"<span style='color:#FFCC00; font-family:Consolas; font-size:14px;'>{char}</span> <span style='color:#999;'>{self.current_status_text}</span>"
        cursor.insertHtml(html_loader)
        self.tab_scan.setTextCursor(cursor)
        self.spin_idx += 1

    def append_final_log(self, html_content):
        cursor = self.tab_scan.textCursor()
        cursor.movePosition(QTextCursor.End)
        cursor.select(QTextCursor.LineUnderCursor)
        cursor.removeSelectedText()
        cursor.insertHtml(html_content)
        cursor.insertBlock()
        self.tab_scan.setTextCursor(cursor)
        self.current_status_text = "Preparando próximo item..."

    def on_process_finished(self, success):
        self.is_processing = False
        self.spinner_timer.stop()
        cursor = self.tab_scan.textCursor()
        cursor.movePosition(QTextCursor.End)
        cursor.select(QTextCursor.LineUnderCursor)
        cursor.removeSelectedText()
        msg = "<br><br><b style='color:#00FFAA; font-size:16px;'>CONCLUÍDO COM SUCESSO.</b>" if success else "<br><br><b style='color:#FF3366; font-size:16px;'>ERRO NO PROCESSAMENTO.</b>"
        self.tab_scan.append(msg)
        sb = self.tab_scan.verticalScrollBar(); sb.setValue(sb.maximum())
        self.sync_ui()

    def sync_ui(self):
        self.table.setRowCount(len(self.live_data))
        yrs, auths, title_words, abstract_words, journals, pubs = [], [], [], [], [], []
        stop = {'a', 'o', 'de', 'do', 'da', 'em', 'um', 'para', 'com', 'no', 'na', 'os', 'as', 'dos', 'das', 'por', 
                'the', 'of', 'and', 'in', 'to', 'for', 'with', 'on', 'an', 'at', 'by', 'is', 'are', 'that', 'this', 'from', 'as',
                'study', 'analysis', 'results', 'data', 'using', 'based', 'paper', 'method', 'model', 'system', 'proposed', 'used'}
        
        abs_count = 0
        for r, row in enumerate(self.live_data):
            for c, val in enumerate(row):
                item = QTableWidgetItem(str(val))
                if c == 3 and str(val).startswith("http"):
                    item.setForeground(QColor("#00D2FF")); font = item.font(); font.setUnderline(True); item.setFont(font)
                if c == 7: 
                    if val != "N/A": abs_count += 1
                    item.setToolTip(str(val)[:500] + "...") 
                self.table.setItem(r, c, item)
            
            auths.extend([a.strip() for a in row[0].split(';') if len(a) > 3])
            if str(row[2]).isdigit(): yrs.append(str(row[2]))
            t_words = re.findall(r'\w+', row[1].lower())
            title_words.extend([w for w in t_words if w not in stop and len(w) > 3])
            if row[7] and row[7] != "N/A":
                a_words = re.findall(r'\w+', row[7].lower())
                abstract_words.extend([w for w in a_words if w not in stop and len(w) > 3])
            if row[4] and row[4] != "N/A": journals.append(row[4])
            if row[5] and row[5] != "N/A": pubs.append(row[5])
        
        self.k1.v.setText(str(len(self.live_data)))
        self.k2.v.setText(str(abs_count))
        self.k4.v.setText(str(len(set(journals))))
        if yrs: self.k3.v.setText(str(max(yrs)))
        
        self.c_yrs.update_data(Counter(yrs))
        self.c_tps.update_data(Counter([row[6] for row in self.live_data]))
        self.c_pub.update_data(Counter(pubs))
        self.b_journal.update_data(Counter(journals))
        self.b_aut.update_data(Counter(auths))
        self.b_words.update_data(Counter(title_words))
        self.b_abstract_words.update_data(Counter(abstract_words))

    def export_bibtex(self):
        path, _ = QFileDialog.getSaveFileName(self, "Exportar BibTeX", "", "*.bib")
        if path and self.live_data:
            with open(path, "w", encoding="utf-8") as f:
                for r in self.live_data:
                    k = re.sub(r'\W+', '', r[0].split(',')[0]) + str(r[2])
                    abs_field = f"  abstract={{{r[7]}}},\n" if r[7] != "N/A" else ""
                    f.write(f"@article{{{k},\n  author={{{r[0]}}},\n  title={{{r[1]}}},\n  journal={{{r[4]}}},\n  year={{{r[2]}}},\n  doi={{{r[3]}}},\n{abs_field}  publisher={{{r[5]}}}\n}}\n\n")

    def export_excel(self):
        path, _ = QFileDialog.getSaveFileName(self, "Exportar Excel", "", "*.xlsx")
        if path and self.live_data:
            df = pd.DataFrame(self.live_data, columns=["Autores", "Título", "Ano", "DOI", "Jornal", "Publisher", "Tipo", "Abstract"])
            df.to_excel(path, index=False)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = BiblioApp()
    w.show()
    sys.exit(app.exec_())