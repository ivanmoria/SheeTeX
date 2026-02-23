import sys, re, webbrowser, math, html
import pandas as pd
import os
import platform
import requests
from collections import Counter

# --- NOVOS IMPORTS PARA IA E PROCESSAMENTO DE TEXTO ---
import nltk
from nltk.corpus import stopwords
from sklearn.feature_extraction.text import TfidfVectorizer

# PyQt5 Imports
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QPushButton, QTextBrowser, QLabel, QFileDialog, 
                             QTabWidget, QTableWidget, QTableWidgetItem, 
                             QHeaderView, QHBoxLayout, QFrame, QScrollArea, QToolTip,
                             QSplitter, QShortcut, QMessageBox)
from PyQt5.QtCore import QThread, pyqtSignal, Qt, QRectF, QTimer, QPoint, QUrl
from PyQt5.QtGui import QPainter, QColor, QFont, QTextCursor, QKeySequence
from PyQt5.QtWebEngineWidgets import QWebEngineView, QWebEngineSettings 

# PDF & Data Imports
from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextContainer, LAParams
from habanero import Crossref

# Inicialização de recursos de IA (Download silencioso)
try:
    nltk.data.find('corpora/stopwords')
except LookupError:
    nltk.download('stopwords', quiet=True)

# =============================================================================
# MOTOR DE IA: EXTRAÇÃO DE TERMOS RELEVANTES (TF-IDF)
# =============================================================================
class ContentAI:
    @staticmethod
    def get_keywords(texts, top_n=10):
        """
        Analisa os abstracts usando TF-IDF para encontrar palavras que 
        realmente definem o conteúdo, ignorando conectivos.
        """
        valid_texts = [t for t in texts if t and len(t) > 30 and t != "N/A"]
        if len(valid_texts) < 2:
            # Fallback caso haja poucos abstracts: contagem simples
            all_words = " ".join(valid_texts).lower().split()
            stops = set(stopwords.words('portuguese') + stopwords.words('english'))
            filtered = [w for w in all_words if len(w) > 3 and w not in stops]
            return dict(Counter(filtered).most_common(top_n))

        try:
            stop_words = list(stopwords.words('portuguese')) + list(stopwords.words('english'))
            # Adiciona termos comuns que não agregam valor acadêmico
            stop_words.extend(['abstract', 'introduction', 'results', 'based', 'study', 'paper', 'using'])
            
            vectorizer = TfidfVectorizer(
                max_features=100,
                stop_words=stop_words,
                ngram_range=(1, 2) # Pega palavras simples e termos compostos (ex: "Inteligência Artificial")
            )
            
            tfidf_matrix = vectorizer.fit_transform(valid_texts)
            sums = tfidf_matrix.sum(axis=0)
            data = []
            for col, term in enumerate(vectorizer.get_feature_names_out()):
                data.append((term, sums[0, col]))
            
            ranking = sorted(data, key=lambda x: x[1], reverse=True)
            return {item[0].title(): int(item[1] * 10) for item in ranking[:top_n]}
        except:
            return {}

# =============================================================================
# ENGINE: PROCESSAMENTO DE PDF (WORKER)
# =============================================================================
class DeepExtractionWorker(QThread):
    initial_count = pyqtSignal(int)
    progress_update = pyqtSignal(int, int)
    item_completed = pyqtSignal(list)
    log_result = pyqtSignal(str)
    finished_signal = pyqtSignal(bool)

    def __init__(self, pdf_path):
        super().__init__()
        self.pdf_path = pdf_path
        self.file_identifier = os.path.basename(pdf_path)

    def clean_text(self, text):
        if not text: return ""
        text = html.unescape(text)
        return re.sub(r'\s+', ' ', text).strip().strip(',').strip('.')

    def clean_abstract(self, text):
        if not text: return "N/A"
        text = html.unescape(text)
        clean = re.sub(r'<[^>]+>', '', text)
        return re.sub(r'\s+', ' ', clean).strip()

    def format_person(self, p):
        fn = p.get('given', '').strip()
        ln = p.get('family', p.get('name', '')).strip() 
        return f"{self.clean_text(ln)}, {self.clean_text(fn)}" if ln and fn else self.clean_text(ln)

    def get_abstract_fallback(self, doi):
        if not doi: return ""
        # 1. SEMANTIC SCHOLAR
        try:
            res = requests.get(f"https://api.semanticscholar.org/graph/v1/paper/DOI:{doi}?fields=abstract", timeout=5) 
            if res.status_code == 200:
                data = res.json()
                if data.get('abstract'): return data['abstract']
        except: pass 

        # 2. OPENALEX
        try:
            headers = {'User-Agent': 'mailto:pesquisador@universidade.edu'}
            res = requests.get(f"https://api.openalex.org/works/https://doi.org/{doi}", headers=headers, timeout=5)
            if res.status_code == 200:
                data = res.json()
                abstract_inverted = data.get('abstract_inverted_index')
                if abstract_inverted: return self.reconstruct_openalex_abstract(abstract_inverted)
        except: pass
        return "" 

    def reconstruct_openalex_abstract(self, inverted_index):
        if not inverted_index: return ""
        word_index = []
        for word, positions in inverted_index.items():
            for pos in positions: word_index.append((pos, word))
        word_index.sort(key=lambda x: x[0])
        return " ".join([word for pos, word in word_index])

    def run(self):
        cr = Crossref()
        try:
            laparams = LAParams(char_margin=3.5)
            full_text = ""; first_page_text = ""; page_count = 0
            
            for page in extract_pages(self.pdf_path, laparams=laparams):
                page_text = ""
                for el in page:
                    if isinstance(el, LTTextContainer): page_text += el.get_text()
                if page_count == 0: first_page_text = page_text
                full_text += page_text; page_count += 1

            # Header do Log
            doc_title = os.path.basename(self.pdf_path)
            self.log_result.emit(f"<div style='color:#00D2FF; font-weight:bold; margin-bottom:10px;'>Iniciando extração: {doc_title}</div>")

            # Identifica referências
            content = full_text
            for kw in ["references", "bibliography", "referências", "bibliografia"]:
                found = full_text.lower().rfind(kw)
                if found != -1: content = full_text[found:]; break
            
            segments = [s.strip().replace('\n', ' ') for s in re.split(r"\n(?=\[?\d+\]?[\.\s]|\b[A-Z][a-z]+, [A-Z]\.)", content) if len(s.strip()) > 15]
            total = len(segments)
            self.initial_count.emit(total)
            history = {} 

            for i, seg in enumerate(segments):
                idx = i + 1
                self.progress_update.emit(idx, total)
                try:
                    doi_m = re.search(r"10\.\d{4,9}/[-._;()/:A-Z0-9]+", seg, re.IGNORECASE)
                    item = None
                    if doi_m:
                        try: item = cr.works(ids=doi_m.group(0).rstrip('.'))['message']
                        except: pass
                    
                    if not item:
                        q = re.sub(r'[^\w\s]', '', seg[:160])
                        res = cr.works(query=q, limit=1)
                        if res['message']['items']: item = res['message']['items'][0]

                    if item:
                        auth_str = "; ".join([self.format_person(a) for a in item.get('author', [])]) or "Autor Desconhecido"
                        year = str(item.get("created", {}).get("date-parts", [[None]])[0][0])
                        title = self.clean_text(item.get("title", ["Unknown Title"])[0])
                        journal = self.clean_text(item.get("container-title", ["N/A"])[0])
                        res_doi = item.get('DOI', '').lower()
                        link = f"https://doi.org/{res_doi}" if res_doi else ""
                        uid = res_doi if res_doi else f"{auth_str}{title}{year}".lower()
                        
                        abstract_raw = item.get('abstract', '')
                        if not abstract_raw and res_doi:
                            abstract_raw = self.get_abstract_fallback(res_doi)
                        abstract_clean = self.clean_abstract(abstract_raw)

                        if uid not in history:
                            history[uid] = title
                            pub = self.clean_text(item.get("publisher", "N/A"))
                            item_type = item.get("type", "other").title()
                            self.item_completed.emit([auth_str, title, year, link, journal, pub, item_type, abstract_clean])
                            self.log_result.emit(f"<div style='color:#BBB; font-size:11px;'>[{idx:03d}] {title[:70]}...</div>")
                except: continue
            
            self.finished_signal.emit(True)
        except Exception as e:
            self.finished_signal.emit(False)

# =============================================================================
# COMPONENTES GRÁFICOS
# =============================================================================
class DonutChart(QWidget):
    def __init__(self, title):
        super().__init__(); self.title, self.data = title, []
        self.setMinimumHeight(320); self.setMouseTracking(True)
        self.colors = [QColor("#00D2FF"), QColor("#7B2FF7"), QColor("#00FFAA"), QColor("#FFCC00"), QColor("#FF3366"), QColor("#FFFFFF")]
    def update_data(self, d): self.data = sorted(d.items(), key=lambda x: x[1], reverse=True)[:6]; self.update()
    def paintEvent(self, event):
        p = QPainter(self); p.setRenderHint(QPainter.Antialiasing)
        p.setPen(QColor("#00D2FF")); p.setFont(QFont("Segoe UI", 12, QFont.Bold))
        p.drawText(0, 0, self.width(), 30, Qt.AlignCenter, self.title.upper())
        if not self.data: return
        total, start = sum(v for k,v in self.data), 90*16
        rect = QRectF(self.width()//2 - 110, 60, 220, 220)
        for i, (l, v) in enumerate(self.data):
            span = -int((v/total)*360*16)
            p.setBrush(self.colors[i%len(self.colors)]); p.setPen(Qt.NoPen); p.drawPie(rect, start, span); start += span
        p.setBrush(QColor("#050505")); p.drawEllipse(self.width()//2 - 60, 110, 120, 120)

class BarChart(QWidget):
    def __init__(self, title, color_hex="#00D2FF"):
        super().__init__(); self.title, self.data, self.rects = title, [], []
        self.bar_color = QColor(color_hex); self.setMinimumHeight(300); self.setMouseTracking(True); self.margin_left = 200
    def update_data(self, d): 
        if isinstance(d, Counter): self.data = d.most_common(8)
        else: self.data = sorted(d.items(), key=lambda x: x[1], reverse=True)[:8]
        self.update()
    def paintEvent(self, event):
        p = QPainter(self); p.setRenderHint(QPainter.Antialiasing)
        p.setPen(QColor("#00D2FF")); p.setFont(QFont("Segoe UI", 12, QFont.Bold))
        p.drawText(self.margin_left, 0, self.width()-self.margin_left, 30, Qt.AlignLeft, self.title.upper())
        if not self.data: return
        max_v = self.data[0][1] if self.data[0][1] > 0 else 1
        start_y, bar_h, gap = 40, 20, 10; self.rects = []
        for l, v in self.data:
            w = int((v/max_v)*(self.width() - self.margin_left - 60))
            r = QRectF(self.margin_left, start_y, w, bar_h); self.rects.append((r, l, v))
            p.setBrush(self.bar_color); p.setPen(Qt.NoPen); p.drawRect(r)
            p.setPen(QColor("#AAA")); p.setFont(QFont("Segoe UI", 9))
            p.drawText(5, int(start_y + bar_h/2 + 4), self.margin_left - 10, bar_h, Qt.AlignRight, str(l)[:25])
            p.setPen(QColor("#FFF")); p.drawText(int(self.margin_left + w + 10), int(start_y + bar_h - 4), str(v))
            start_y += (bar_h + gap)

# =============================================================================
# INTERFACE PRINCIPAL
# =============================================================================
class BiblioApp(QMainWindow):
    def __init__(self):
        super().__init__(); 
        self.live_data = [] 
        self.initUI()
        self.spinner_timer = QTimer()
        self.spinner_timer.timeout.connect(self.update_progress_label)

    def initUI(self):
        self.setWindowTitle("BiblioTech AI Dashboard"); self.resize(1600, 1000)
        self.setStyleSheet("""
            QMainWindow, QWidget { background: #050505; color: #FFF; font-family: 'Segoe UI', Arial; }
            QTableWidget { background: #020202; gridline-color: #111; color: #CCC; border: none; font-size: 13px; }
            QHeaderView::section { background: #0A0A0A; color: #00D2FF; font-weight: bold; padding: 12px; border: none; }
            QPushButton { background: #111; border: 1px solid #333; color: #BBB; padding: 6px 15px; font-weight: bold; font-size: 11px; min-width: 100px; border-radius: 4px; }
            QPushButton:hover { background: #222; color: #FFF; border-color: #555; }
            QPushButton#ActionBtn { background: #003344; color: #00D2FF; border: 1px solid #005577; }
            QTabWidget::pane { border-top: 1px solid #151515; } 
            QTabBar::tab { background: #0A0A0A; padding: 10px 20px; color: #666; font-weight: bold; }
            QTabBar::tab:selected { color: #FFF; border-bottom: 2px solid #00D2FF; background: #0F0F0F; }
        """)

        self.tabs = QTabWidget(); self.setCentralWidget(self.tabs)

        # Toolbar superior
        self.left_widget = QWidget(); self.left_layout = QHBoxLayout(self.left_widget)
        self.btn_load = QPushButton("+ ADICIONAR PDF"); self.btn_load.setObjectName("ActionBtn")
        self.btn_load.clicked.connect(self.open_pdf)
        self.lbl_count = QLabel("0 Refs Carregadas"); self.lbl_count.setStyleSheet("color: #666; margin-left:10px;")
        self.left_layout.addWidget(self.btn_load); self.left_layout.addWidget(self.lbl_count)
        self.tabs.setCornerWidget(self.left_widget, Qt.TopLeftCorner)

        # Aba 1: Log/PDF
        self.setup_tab1()
        # Aba 2: Tabela
        self.setup_tab2()
        # Aba 3: Dashboard
        self.setup_tab3()

    def setup_tab1(self):
        w = QWidget(); l = QVBoxLayout(w); l.setContentsMargins(0,0,0,0)
        self.splitter = QSplitter(Qt.Horizontal)
        self.tab_scan = QTextBrowser(); self.tab_scan.setStyleSheet("background:#000; border:none;")
        self.pdf_viewer = QWebEngineView(); self.pdf_viewer.setStyleSheet("background:#050505;")
        self.splitter.addWidget(self.tab_scan); self.splitter.addWidget(self.pdf_viewer)
        self.progress_label = QLabel(""); self.progress_label.setFixedHeight(16)
        l.addWidget(self.splitter); l.addWidget(self.progress_label)
        self.tabs.addTab(w, "SCANNER LOG")

    def setup_tab2(self):
        w = QWidget(); l = QVBoxLayout(w)
        self.table = QTableWidget(0, 9)
        self.table.setHorizontalHeaderLabels(["ARQUIVO", "AUTORES", "TÍTULO", "ANO", "DOI", "JORNAL", "PUBLISHER", "TIPO", "ABSTRACT"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table.cellClicked.connect(self.handle_cell_click)
        l.addWidget(self.table)
        self.tabs.addTab(w, "DATABASE")

    def setup_tab3(self):
        scroll = QScrollArea(); scroll.setWidgetResizable(True); scroll.setStyleSheet("border:none;")
        w = QWidget(); self.l_dash = QVBoxLayout(w); self.setup_dash_widgets()
        scroll.setWidget(w); self.tabs.addTab(scroll, "DASHBOARD ANALÍTICO")

    def setup_dash_widgets(self):
        self.l_dash.setContentsMargins(40,40,40,40); self.l_dash.setSpacing(30)
        
        # KPIs
        h_kpi = QHBoxLayout()
        self.k1 = self.create_kpi("TOTAL REFERÊNCIAS")
        self.k2 = self.create_kpi("ABSTRACTS (IA)")
        self.k3 = self.create_kpi("TERMOS ÚNICOS") # IA KPI
        self.k4 = self.create_kpi("FONTES")
        h_kpi.addWidget(self.k1); h_kpi.addWidget(self.k2); h_kpi.addWidget(self.k4); h_kpi.addWidget(self.k3)
        self.l_dash.addLayout(h_kpi)

        # Gráficos IA e Pizza
        h_row2 = QHBoxLayout()
        self.c_yrs = DonutChart("Por Ano")
        self.b_keywords = BarChart("Palavras Mais Repetidas (IA TF-IDF)", "#00FFAA") # Lógica pedida
        h_row2.addWidget(self.c_yrs); h_row2.addWidget(self.b_keywords)
        self.l_dash.addLayout(h_row2)

        self.btn_sync = QPushButton("RECALCULAR ANALYTICS COM IA")
        self.btn_sync.setObjectName("ActionBtn")
        self.btn_sync.clicked.connect(self.update_dashboard_view)
        self.l_dash.addWidget(self.btn_sync)

    def create_kpi(self, t):
        f = QFrame(); f.setStyleSheet("background: #0F0F0F; border: 1px solid #222; border-radius: 12px;"); f.setFixedHeight(110)
        l = QVBoxLayout(f); title = QLabel(t); title.setStyleSheet("color:#888; font-size:12px;")
        v = QLabel("-"); v.setStyleSheet("color:#FFF; font-size:32px; font-weight:900;")
        title.setAlignment(Qt.AlignCenter); v.setAlignment(Qt.AlignCenter)
        l.addWidget(title); l.addWidget(v); f.v = v; return f

    def open_pdf(self):
        path, _ = QFileDialog.getOpenFileName(self, "Abrir PDF", "", "PDF (*.pdf)")
        if path:
            self.pdf_viewer.load(QUrl.fromLocalFile(path))
            self.worker = DeepExtractionWorker(path)
            self.worker.item_completed.connect(self.add_item)
            self.worker.log_result.connect(lambda m: self.tab_scan.append(m))
            self.worker.start()

    def add_item(self, data):
        self.live_data.append(data)
        self.lbl_count.setText(f"{len(self.live_data)} Refs Carregadas")
        r = self.table.rowCount(); self.table.insertRow(r)
        # Adiciona o nome do arquivo na primeira coluna (índice 0)
        self.table.setItem(r, 0, QTableWidgetItem(os.path.basename(self.worker.pdf_path)))
        for i, val in enumerate(data):
            self.table.setItem(r, i+1, QTableWidgetItem(str(val)))

    def update_dashboard_view(self):
        if not self.live_data: return
        # Auth, Title, Year, Link, Journal, Pub, Type, Abstract
        df = pd.DataFrame(self.live_data, columns=["Auth", "Title", "Year", "Link", "Journal", "Pub", "Type", "Abstract"])
        
        # 1. KPIs
        self.k1.v.setText(str(len(df)))
        abstracts = [a for a in df['Abstract'].tolist() if a != "N/A"]
        self.k2.v.setText(str(len(abstracts)))
        self.k4.v.setText(str(df['Journal'].nunique()))

        # 2. IA Keywords (Lógica TF-IDF)
        keywords = ContentAI.get_keywords(abstracts)
        self.k3.v.setText(str(len(keywords)))
        self.b_keywords.update_data(keywords)

        # 3. Anos
        self.c_yrs.update_data(dict(df['Year'].value_counts()))

    def handle_cell_click(self, row, col):
        data = {
            "titulo": self.table.item(row, 2).text(),
            "autores": self.table.item(row, 1).text(),
            "jornal": self.table.item(row, 5).text(),
            "ano": self.table.item(row, 3).text(),
            "doi": self.table.item(row, 4).text(),
            "abstract": self.table.item(row, 8).text()
        }
        self.show_abstract_popup(data)

    def show_abstract_popup(self, data):
        msg = QMessageBox(self); msg.setWindowTitle("Detalhes da Referência")
        tb = QTextBrowser(); tb.setOpenExternalLinks(True); tb.setMinimumSize(700, 500)
        html = f"""
        <div style='font-family:sans-serif; background:#111; padding:20px; color:#EEE;'>
            <h2 style='color:#00D2FF;'>{data['titulo']}</h2>
            <p><b>Autores:</b> {data['autores']}<br>
            <b>Fonte:</b> {data['jornal']} ({data['ano']})</p>
            <hr style='border:1px solid #333;'>
            <div style='margin-top:20px; line-height:1.6;'><b>Resumo:</b><br>{data['abstract']}</div>
        </div>
        """
        tb.setHtml(html)
        msg.layout().addWidget(tb, 0, 0, 1, msg.layout().columnCount())
        msg.exec_()

    def update_progress_label(self): pass # Implementação opcional
    def export_bibtex(self): pass
    def export_excel(self): pass
    def update_table_view(self): pass

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = BiblioApp()
    window.show()
    sys.exit(app.exec_())