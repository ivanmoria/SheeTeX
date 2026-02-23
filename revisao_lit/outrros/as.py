import sys, re, webbrowser, math, html
import pandas as pd
import os
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from collections import Counter
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QPushButton, QTextBrowser, QLabel, QFileDialog, 
                             QTabWidget, QTableWidget, QTableWidgetItem, 
                             QHeaderView, QHBoxLayout, QFrame, QScrollArea, QToolTip)
from PyQt5.QtCore import QThread, pyqtSignal, Qt, QRectF, QTimer, QPoint
from PyQt5.QtGui import QPainter, QColor, QFont, QTextCursor
from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextContainer, LAParams
from habanero import Crossref

# =============================================================================
# ENGINE: PROCESSAMENTO DE PDF (WORKER)
# =============================================================================
class DeepExtractionWorker(QThread):
    # Signals
    initial_count = pyqtSignal(int)
    status_update = pyqtSignal(str) # Texto curto para o spinner (ex: "Buscando...")
    item_completed = pyqtSignal(list)
    log_result = pyqtSignal(str)    # HTML formatado do item concluído
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
                
                # Extração Única
                for page in extract_pages(self.pdf_path, laparams=laparams):
                    for el in page:
                        if isinstance(el, LTTextContainer): 
                            full_text += el.get_text()

                # --- 1. CABEÇALHO DO DOCUMENTO (Extraído do início do texto) ---
                file_name = os.path.basename(self.pdf_path)
                lines = [l.strip() for l in full_text.split('\n') if len(l.strip()) > 3]
                
                # Tentamos pegar o Título (1ª linha válida) e Autor (2ª linha válida)
                doc_title = lines[0] if len(lines) > 0 else file_name
                doc_author = lines[1] if len(lines) > 1 else "Unknown Author"
                
                # Tenta achar um ano (4 dígitos começando com 19 ou 20) no início do texto
                year_match = re.search(r"\b(19|20)\d{2}\b", full_text[:1000])
                doc_year = year_match.group(0) if year_match else "n/a"

                header_html = (
                    f"<div style='background-color: #1a1a1a; padding: 8px; border-left: 4px solid #00D2FF; margin-bottom: 10px;'>"
                    f"  <b style='color: #00D2FF; font-size: 10px; font-family: sans-serif;'>PROCESSING FILE:</b><br>"
                    f"  <b style='color: #FFF; font-size: 14px;'>{doc_title[:150]}</b><br>"
                    f"  <span style='color: #AAA; font-size: 11px;'>{doc_author[:100]} • {doc_year}</span>"
                    f"</div>"
                )
                self.log_result.emit(header_html)

                # --- 2. SEGMENTAÇÃO DAS REFERÊNCIAS ---
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
                    self.status_update.emit(f"Processing Crossref {idx}/{total}...")
                    
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
                            
                            is_duplicate = uid in history
                            
                            if is_duplicate:
                                idx_color = "#FF3366" 
                                msg_dup = "<span style='color:#FF3366; font-weight:bold; font-size:10px;'> [Duplicate]</span>"
                                main_color = "#555"
                                sub_color = "#444"
                                doi_style = "color:#444; text-decoration:none;"
                            else:
                                idx_color = "#00D2FF" 
                                msg_dup = ""
                                main_color = "#FFF"
                                sub_color = "#CCC"
                                # Link azul e sublinhado para simular link clicável/hover
                                doi_style = "color:#00D2FF; text-decoration:underline;"

                            doi_html = f"<a href='{link}' style='{doi_style}'>{res_doi}</a>" if link else "N/A"
                            
                            detail_log = (
                                f"<div style='margin-bottom: 2px; border-bottom: 1px solid #222; padding-bottom: 2px; line-height: 1.0;'>"
                                f"  <span style='color:{idx_color}; font-family:Consolas; font-weight:bold; font-size:12px;'>[{idx:03d}/{total:03d}]</span>{msg_dup} "
                                f"  <b style='color:{main_color}; font-size:13px;'>{title}</b><br>"
                                f"  <span style='color:{sub_color}; font-size:11px;'>{auth_str}</span><br>"
                                f"  <span style='color:#666; font-size:11px;'>{journal} • {year} • DOI: {doi_html}</span>"
                                f"</div>"
                            )

                            if not is_duplicate:
                                history[uid] = title
                                pub = self.clean_text(item.get("publisher", "N/A"))
                                self.item_completed.emit([auth_str, title, year, link, journal, pub, item.get("type", "other").title()])
                            
                            self.log_result.emit(detail_log)
                        else:
                            self.log_result.emit(f"<div style='color:#FF3366; font-size:11px; margin-bottom: 2px;'>[{idx:03d}/{total:03d}] Reference not identified.</div>")
                    
                    except Exception as e:
                        self.log_result.emit(f"<div style='color:#FF3366; font-size:11px; margin-bottom: 2px;'>[{idx:03d}/{total:03d}] Error: {str(e)}</div>")
                
                self.finished_signal.emit(True)
            except Exception as e: 
                self.finished_signal.emit(False)

# =============================================================================
# COMPONENTES VISUAIS (GRÁFICOS)
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
            p.drawPie(rect, start, span)
            start += span
            
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
        max_v = self.data[0][1]
        start_y, bar_h, gap = 40, 20, 10
        self.rects = []
        
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
# INTERFACE PRINCIPAL
# =============================================================================
class BiblioApp(QMainWindow):
    def __init__(self):
        super().__init__(); self.live_data = []
        
        # Variáveis de Estado de Loading
        self.is_processing = False
        self.current_status_text = "Inicializando..."
        self.spinner_chars = ["⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏"]
        self.spin_idx = 0
        
        self.initUI()
        
        # Timer independente para a animação do Spinner
        self.spinner_timer = QTimer()
        self.spinner_timer.timeout.connect(self.update_spinner_animation)

    def initUI(self):
        self.setWindowTitle("BiblioTech Ultimate V24.1"); self.resize(1600, 1000)
        self.setStyleSheet("""
            QMainWindow, QWidget { background: #050505; color: #FFF; font-family: 'Segoe UI', Arial; }
            QTableWidget { background: #020202; gridline-color: #111; color: #CCC; border: none; font-size: 13px; }
            QHeaderView::section { background: #0A0A0A; color: #00D2FF; font-weight: bold; padding: 12px; border: none; }
            QPushButton { 
                background: transparent; border: none; padding: 8px 20px; 
                font-weight: bold; font-size: 12px; min-width: 120px; 
            }
            QPushButton:hover { background: #111; border-radius: 4px; }
            QTabWidget::pane { border-top: 1px solid #151515; top: -1px; } 
            QTabBar::tab { background: #0A0A0A; padding: 12px 30px; color: #444; font-weight: bold; border-bottom: 2px solid transparent; }
            QTabBar::tab:selected { color: #FFF; border-bottom: 2px solid #00D2FF; }
            QScrollBar:vertical { background: #050505; width: 12px; }
            QScrollBar::handle:vertical { background: #333; min-height: 20px; border-radius: 6px; }
        """)

        self.tabs = QTabWidget()
        self.tabs.setDocumentMode(True)
        self.tabs.currentChanged.connect(self.toggle_buttons)
        self.setCentralWidget(self.tabs)

        # Toolbar
        self.left_widget = QWidget()
        self.left_layout = QHBoxLayout(self.left_widget)
        self.left_layout.setContentsMargins(10, 5, 0, 5)
        self.btn_load = QPushButton("IMPORTAR PDF")
        self.btn_load.setStyleSheet("color: #00D2FF; border: 1px solid #003344; border-radius: 4px;")
        self.btn_load.clicked.connect(self.open_pdf)
        self.left_layout.addWidget(self.btn_load)
        self.tabs.setCornerWidget(self.left_widget, Qt.TopLeftCorner)

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

        # Tabs Content
        self.tab_scan = QTextBrowser() 
        self.tab_scan.setOpenExternalLinks(True) 
        self.tab_scan.setStyleSheet("background: #000; border: none; padding: 25px; font-family: 'Segoe UI'; font-size: 13px;")
        self.table = QTableWidget(0, 7)
        self.table.setHorizontalHeaderLabels(["AUTORES", "TÍTULO", "ANO", "LINK DOI", "JORNAL", "PUBLISHER", "TIPO"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table.horizontalHeader().resizeSection(1, 400)
        self.table.cellClicked.connect(self.handle_cell_click)
        
        dash_scroll = QScrollArea(); dash_scroll.setWidgetResizable(True); dash_scroll.setStyleSheet("border: none;") 
        self.dash_inner = QWidget(); self.l_dash = QVBoxLayout(self.dash_inner); self.setup_dash_widgets()
        dash_scroll.setWidget(self.dash_inner)
        
        self.tabs.addTab(self.tab_scan, "SCANNER LOG")
        self.tabs.addTab(self.table, "BANCO DE DADOS")
        self.tabs.addTab(dash_scroll, "DASHBOARD")
        self.toggle_buttons(0)

    def setup_dash_widgets(self):
        self.l_dash.setContentsMargins(40, 40, 40, 40); self.l_dash.setSpacing(40)
        h_kpi = QHBoxLayout(); h_kpi.setSpacing(20)
        self.k1 = self.create_kpi("TOTAL")
        self.k2 = self.create_kpi("AUTORES")
        self.k3 = self.create_kpi("ANO REC.")
        self.k4 = self.create_kpi("JORNAIS")
        h_kpi.addWidget(self.k1); h_kpi.addWidget(self.k2); h_kpi.addWidget(self.k4); h_kpi.addWidget(self.k3)
        self.l_dash.addLayout(h_kpi)

        h_donuts = QHBoxLayout(); h_donuts.setSpacing(20)
        self.c_yrs = DonutChart("Por Ano")
        self.c_tps = DonutChart("Tipos")
        self.c_pub = DonutChart("Top Publishers")
        h_donuts.addWidget(self.c_yrs); h_donuts.addWidget(self.c_tps); h_donuts.addWidget(self.c_pub)
        self.l_dash.addLayout(h_donuts)

        h_bars = QHBoxLayout(); h_bars.setSpacing(20)
        self.b_journal = BarChart("Top Jornais / Conferências", "#FF3366")
        self.b_aut = BarChart("Autores Mais Citados", "#7B2FF7")
        h_bars.addWidget(self.b_journal); h_bars.addWidget(self.b_aut)
        self.l_dash.addLayout(h_bars)
        
        self.b_words = BarChart("Termos Frequentes nos Títulos", "#00FFAA")
        self.l_dash.addWidget(self.b_words)

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
        if col == 3:
            url = self.table.item(row, col).text()
            if url.startswith("http"): webbrowser.open(url)

    def open_pdf(self):
        p, _ = QFileDialog.getOpenFileName(self, "Selecionar PDF", "", "*.pdf")
        if p:
            self.tab_scan.clear(); self.live_data = []
            self.tabs.setCurrentIndex(0)
            
            # Prepara o worker
            self.worker = DeepExtractionWorker(p)
            self.worker.initial_count.connect(self.handle_initial_count) # Sinal para o cabeçalho
            self.worker.status_update.connect(self.update_status_text)
            self.worker.item_completed.connect(self.live_data.append)
            self.worker.log_result.connect(self.append_final_log)
            self.worker.finished_signal.connect(self.on_process_finished)
            
            self.worker.start()

    # --- LÓGICA DO LOG E SPINNER ---
    
    def handle_initial_count(self, total):
        """Imprime o cabeçalho inicial"""
        header = (
            f"<div style='margin-bottom: 20px; font-size: 14px; color: #FFF;'>"
            f"<b>SISTEMA INICIADO</b><br>"
            f"<span style='color: #00D2FF;'>Encontradas {total} referências bibliográficas.</span><br>"

            f"</div><br>" 
        )
        cursor = self.tab_scan.textCursor()
        cursor.insertHtml(header)
        # Garante que o cursor esteja numa nova linha para o spinner começar
        cursor.movePosition(QTextCursor.End)
        self.tab_scan.setTextCursor(cursor)
        
        # Inicia a animação apenas agora
        self.is_processing = True
        self.spinner_timer.start(80)

    def update_status_text(self, text):
        self.current_status_text = text

    def update_spinner_animation(self):
        if not self.is_processing: return

        cursor = self.tab_scan.textCursor()
        cursor.movePosition(QTextCursor.End)
        
        # Seleciona a linha atual (onde o spinner está rodando)
        cursor.select(QTextCursor.LineUnderCursor)
        cursor.removeSelectedText()
        
        char = self.spinner_chars[self.spin_idx % len(self.spinner_chars)]
        html_loader = f"<span style='color:#FFCC00; font-family:Consolas; font-size:14px;'>{char}</span> <span style='color:#999;'>{self.current_status_text}</span>"
        
        cursor.insertHtml(html_loader)
        self.tab_scan.setTextCursor(cursor)
        self.spin_idx += 1

    def append_final_log(self, html_content):
        # 1. Remove a linha do spinner atual
        cursor = self.tab_scan.textCursor()
        cursor.movePosition(QTextCursor.End)
        cursor.select(QTextCursor.LineUnderCursor)
        cursor.removeSelectedText()
        
        # 2. Insere o log definitivo e quebra linha
        cursor.insertHtml(html_content)
        cursor.insertBlock() # Nova linha limpa para o próximo spinner
        
        self.tab_scan.setTextCursor(cursor)
        self.current_status_text = "Preparando próximo item..."

    def on_process_finished(self, success):
        self.is_processing = False
        self.spinner_timer.stop()
        
        # Limpa a última linha de spinner (que ficou "Preparando próximo item...")
        cursor = self.tab_scan.textCursor()
        cursor.movePosition(QTextCursor.End)
        cursor.select(QTextCursor.LineUnderCursor)
        cursor.removeSelectedText()
        
        msg = "<br><br><b style='color:#00FFAA; font-size:16px;'>CONCLUÍDO COM SUCESSO.</b>" if success else "<br><br><b style='color:#FF3366; font-size:16px;'>ERRO NO PROCESSAMENTO.</b>"
        self.tab_scan.append(msg)
        
        # Rola até o final
        sb = self.tab_scan.verticalScrollBar()
        sb.setValue(sb.maximum())
        
        self.sync_ui()

    def sync_ui(self):
        self.table.setRowCount(len(self.live_data))
        yrs, auths, words, journals, pubs = [], [], [], [], []
        stop = {'a', 'o', 'de', 'do', 'da', 'em', 'um', 'para', 'com', 'no', 'na', 'the', 'of', 'and', 'in', 'to', 'for', 'with', 'on', 'an', 'at', 'by', 'study', 'analysis'}
        
        for r, row in enumerate(self.live_data):
            for c, val in enumerate(row):
                item = QTableWidgetItem(str(val))
                if c == 3 and str(val).startswith("http"):
                    item.setForeground(QColor("#00D2FF")); font = item.font(); font.setUnderline(True); item.setFont(font)
                if c == 1: item.setToolTip(str(val))
                self.table.setItem(r, c, item)
            
            auths.extend([a.strip() for a in row[0].split(';') if len(a) > 3])
            if str(row[2]).isdigit(): yrs.append(str(row[2]))
            t_words = re.findall(r'\w+', row[1].lower())
            words.extend([w for w in t_words if w not in stop and len(w) > 3])
            if row[4] and row[4] != "N/A": journals.append(row[4])
            if row[5] and row[5] != "N/A": pubs.append(row[5])
        
        self.k1.v.setText(str(len(self.live_data)))
        self.k2.v.setText(str(len(set(auths))))
        self.k4.v.setText(str(len(set(journals))))
        if yrs: self.k3.v.setText(str(max(yrs)))
        
        self.c_yrs.update_data(Counter(yrs))
        self.c_tps.update_data(Counter([row[6] for row in self.live_data]))
        self.c_pub.update_data(Counter(pubs))
        self.b_journal.update_data(Counter(journals))
        self.b_aut.update_data(Counter(auths))
        self.b_words.update_data(Counter(words))

    def export_bibtex(self):
        path, _ = QFileDialog.getSaveFileName(self, "Exportar BibTeX", "", "*.bib")
        if path and self.live_data:
            with open(path, "w", encoding="utf-8") as f:
                for r in self.live_data:
                    k = re.sub(r'\W+', '', r[0].split(',')[0]) + str(r[2])
                    f.write(f"@article{{{k},\n  author={{{r[0]}}},\n  title={{{r[1]}}},\n  journal={{{r[4]}}},\n  year={{{r[2]}}},\n  doi={{{r[3]}}},\n  publisher={{{r[5]}}}\n}}\n\n")

    def export_excel(self):
        path, _ = QFileDialog.getSaveFileName(self, "Exportar Excel", "", "*.xlsx")
        if path and self.live_data:
            df = pd.DataFrame(self.live_data, columns=["Autores", "Título", "Ano", "DOI", "Jornal", "Publisher", "Tipo"])
            df.to_excel(path, index=False)

if __name__ == "__main__":
    app = QApplication(sys.argv); w = BiblioApp(); w.show(); sys.exit(app.exec_())