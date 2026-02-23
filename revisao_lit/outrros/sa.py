import sys, re, webbrowser, math, html
import pandas as pd
import os
from collections import Counter

# PyQt5 Imports
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QPushButton, QTextBrowser, QLabel, QFileDialog, 
                             QTabWidget, QTableWidget, QTableWidgetItem, 
                             QHeaderView, QHBoxLayout, QFrame, QScrollArea, QToolTip,
                             QSplitter)
from PyQt5.QtCore import QThread, pyqtSignal, Qt, QRectF, QTimer, QPoint, QUrl
from PyQt5.QtGui import QPainter, QColor, QFont, QTextCursor
from PyQt5.QtWebEngineWidgets import QWebEngineView, QWebEngineSettings 

# PDF & Data Imports
from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextContainer, LAParams
from habanero import Crossref

# =============================================================================
# ENGINE: PROCESSAMENTO DE PDF (WORKER)
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
# CLASSES DE GRÁFICOS
# =============================================================================
class DonutChart(QWidget):
    def __init__(self, title):
        super().__init__(); self.title, self.data = title, []
        self.setMinimumHeight(300)
        self.colors = [QColor("#00D2FF"), QColor("#7B2FF7"), QColor("#00FFAA"), QColor("#FFCC00"), QColor("#FF3366")]
    
    def update_data(self, d): 
        self.data = sorted(d.items(), key=lambda x: x[1], reverse=True)[:5]
        self.update()

    def paintEvent(self, event):
        p = QPainter(self); p.setRenderHint(QPainter.Antialiasing)
        p.setPen(QColor("#00D2FF")); p.setFont(QFont("Segoe UI", 10, QFont.Bold))
        p.drawText(self.rect(), Qt.AlignTop | Qt.AlignHCenter, self.title.upper())
        if not self.data: return
        rect = QRectF(self.width()//2-80, self.height()//2-80, 160, 160)
        total, start = sum(v for k,v in self.data), 90*16
        for i, (l, v) in enumerate(self.data):
            span = -int((v/total)*360*16)
            p.setBrush(self.colors[i%5]); p.setPen(Qt.NoPen)
            p.drawPie(rect, start, span); start += span
        p.setBrush(QColor("#050505")); p.drawEllipse(self.width()//2-50, self.height()//2-50, 100, 100)

class BarChart(QWidget):
    def __init__(self, title, color="#00D2FF"):
        super().__init__(); self.title, self.data, self.color = title, [], QColor(color)
        self.setMinimumHeight(250)

    def update_data(self, d): self.data = d.most_common(5); self.update()
        
    def paintEvent(self, event):
        p = QPainter(self); p.setRenderHint(QPainter.Antialiasing)
        p.setPen(QColor("#BBB")); p.drawText(10, 20, self.title)
        if not self.data: return
        max_v = self.data[0][1]; y = 40
        for l, v in self.data:
            w = int((v/max_v)*(self.width()-150))
            p.setBrush(self.color); p.drawRect(120, y, w, 15)
            p.drawText(0, y+12, 110, 15, Qt.AlignRight, str(l)[:15])
            y += 25

# =============================================================================
# INTERFACE PRINCIPAL
# =============================================================================
class BiblioApp(QMainWindow):
    def __init__(self):
        super().__init__(); self.live_data = []
        self.is_processing = False
        self.current_status_text = "Aguardando PDF..."
        self.spinner_chars = ["⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏"]
        self.spin_idx = 0
        self.initUI()
        self.spinner_timer = QTimer()
        self.spinner_timer.timeout.connect(self.update_spinner_animation)

    def initUI(self):
        self.setWindowTitle("BiblioTech Ultimate V2"); self.resize(1500, 900)
        self.setStyleSheet("QMainWindow, QWidget { background: #050505; color: #FFF; font-family: 'Segoe UI'; }")

        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)

        # Botões de Controle
        self.btn_load = QPushButton("IMPORTAR PDF")
        self.btn_load.clicked.connect(self.open_pdf)
        self.btn_load.setStyleSheet("background: #003344; color: #00D2FF; font-weight: bold; padding: 10px;")
        self.tabs.setCornerWidget(self.btn_load, Qt.TopLeftCorner)

        # ABA 1: SCANNER + PDF
        self.tab1_widget = QWidget()
        layout1 = QHBoxLayout(self.tab1_widget)
        self.splitter = QSplitter(Qt.Horizontal)

        self.tab_scan = QTextBrowser()
        self.tab_scan.setStyleSheet("background: #000; border: none; padding: 20px;")
        
        # O LEITOR DE PDF
        self.pdf_viewer = QWebEngineView()
        self.pdf_viewer.settings().setAttribute(QWebEngineSettings.PluginsEnabled, True)
        self.pdf_viewer.settings().setAttribute(QWebEngineSettings.PdfViewerEnabled, True)
        self.pdf_viewer.settings().setAttribute(QWebEngineSettings.LocalContentCanAccessFileUrls, True)
        self.pdf_viewer.setMinimumWidth(500) # Garante que ele apareça

        self.splitter.addWidget(self.tab_scan)
        self.splitter.addWidget(self.pdf_viewer)
        self.splitter.setSizes([700, 800]) # Força a divisão inicial
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
        h1 = QHBoxLayout()
        self.c_yrs = DonutChart("Anos")
        self.c_tps = DonutChart("Tipos")
        h1.addWidget(self.c_yrs); h1.addWidget(self.c_tps)
        
        self.b_journal = BarChart("Principais Jornais", "#FF3366")
        self.b_aut = BarChart("Principais Autores", "#7B2FF7")
        
        self.l_dash.addLayout(h1)
        self.l_dash.addWidget(self.b_journal)
        self.l_dash.addWidget(self.b_aut)

    def open_pdf(self):
        p, _ = QFileDialog.getOpenFileName(self, "Selecionar PDF", "", "*.pdf")
        if p:
            # Garante que o PDF seja carregado via URL local absoluta
            self.pdf_viewer.setUrl(QUrl.fromLocalFile(os.path.abspath(p)))
            self.splitter.setSizes([700, 800]) # Garante que o splitter abra o PDF
            
            self.tab_scan.clear(); self.live_data = []
            self.worker = DeepExtractionWorker(p)
            self.worker.initial_count.connect(lambda t: self.start_process(t))
            self.worker.status_update.connect(lambda s: setattr(self, 'current_status_text', s))
            self.worker.item_completed.connect(self.live_data.append)
            self.worker.log_result.connect(self.tab_scan.append)
            self.worker.finished_signal.connect(self.on_finished)
            self.worker.start()

    def start_process(self, total):
        self.is_processing = True
        self.spinner_timer.start(80)

    def update_spinner_animation(self):
        if not self.is_processing: return
        cursor = self.tab_scan.textCursor()
        cursor.movePosition(QTextCursor.End)
        cursor.select(QTextCursor.LineUnderCursor)
        cursor.removeSelectedText()
        char = self.spinner_chars[self.spin_idx % 10]
        cursor.insertHtml(f"<span style='color:#FFCC00;'>{char}</span> <span style='color:#777;'>{self.current_status_text}</span>")
        self.spin_idx += 1

    def on_finished(self):
        self.is_processing = False
        self.spinner_timer.stop()
        self.sync_data()

    def sync_data(self):
        self.table.setRowCount(len(self.live_data))
        yrs, journals, auths = [], [], []
        for r, row in enumerate(self.live_data):
            for c, val in enumerate(row):
                self.table.setItem(r, c, QTableWidgetItem(str(val)))
            yrs.append(row[2]); journals.append(row[4])
            auths.extend(row[0].split(';'))
        
        self.c_yrs.update_data(Counter(yrs))
        self.b_journal.update_data(Counter(journals))
        self.b_aut.update_data(Counter([a.strip() for a in auths if len(a) > 2]))

if __name__ == "__main__":
    app = QApplication(sys.argv)
    # Estilo Fusion ajuda na compatibilidade do WebEngine no Windows
    app.setStyle("Fusion") 
    w = BiblioApp()
    w.show()
    sys.exit(app.exec_())