import sys, re, webbrowser, math, html
import pandas as pd
import os
import platform
import requests # Import adicionado para as chamadas de API
from collections import Counter

# PyQt5 Imports
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QPushButton, QTextBrowser, QLabel, QFileDialog, QProgressBar,
                             QTabWidget, QTableWidget, QTableWidgetItem, QSizePolicy,
                             QHeaderView, QHBoxLayout, QFrame, QScrollArea, QToolTip,
                             QSplitter, QShortcut, QMessageBox, QDialog, QFormLayout, QLineEdit, QTextEdit, QDialogButtonBox)
from PyQt5.QtCore import QThread, pyqtSignal, Qt, QRectF, QTimer, QPoint, QUrl
from PyQt5.QtGui import QPainter, QColor, QFont, QTextCursor, QKeySequence
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
    progress_update = pyqtSignal(int, int)  # índice atual, total
    item_completed = pyqtSignal(list)       # dados da referência
    log_result = pyqtSignal(str)             # HTML para o log principal
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

    # --- NOVOS MÉTODOS DE EXTRAÇÃO INTELIGENTE DE ABSTRACTS ---
    def get_abstract_fallback(self, doi):
        """
        Tenta buscar o abstract em APIs alternativas abertas quando o Crossref falha.
        """
        if not doi:
            return ""
        try:
            s2_url = f"https://api.semanticscholar.org/graph/v1/paper/DOI:{doi}?fields=abstract"
            s2_res = requests.get(s2_url, timeout=5) 
            if s2_res.status_code == 200:
                data = s2_res.json()
                if data.get('abstract'):
                    return data['abstract']
        except Exception:
            pass 
        try:
            oa_url = f"https://api.openalex.org/works/https://doi.org/{doi}"
            headers = {'User-Agent': 'mailto:pesquisador@universidade.edu'}
            oa_res = requests.get(oa_url, headers=headers, timeout=5)
            if oa_res.status_code == 200:
                data = oa_res.json()
                abstract_inverted = data.get('abstract_inverted_index')
                if abstract_inverted:
                    return self.reconstruct_openalex_abstract(abstract_inverted)
        except Exception:
            pass
        return "" 

    def reconstruct_openalex_abstract(self, inverted_index):
        if not inverted_index: return ""
        word_index = []
        for word, positions in inverted_index.items():
            for pos in positions:
                word_index.append((pos, word))
        word_index.sort(key=lambda x: x[0])
        return " ".join([word for pos, word in word_index])

    def run(self):
        cr = Crossref()
        try:
            # Mensagens de progresso iniciais
            self.progress_update.emit(0, 0)  

            laparams = LAParams(char_margin=3.5)
            full_text = ""
            first_page_text = ""
            
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

            doc_title = os.path.basename(self.pdf_path)
            doc_author = "Autor Desconhecido"
            doc_year = "n.d."
            doc_journal = "Periódico Não Identificado"
            doc_doi_link = ""
            doc_full_ref = "Referência completa não disponível."

            main_doi_match = re.search(r"10\.\d{4,9}/[-._;()/:A-Z0-9]+", first_page_text, re.IGNORECASE)
            main_item = None
            if main_doi_match:
                try: main_item = cr.works(ids=main_doi_match.group(0).rstrip('.'))['message']
                except: pass
            
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

            # --- 3. LOG HEADER ---
            header_html = (
                f"<div style='padding: 15px 0; border-bottom: 2px solid #333; margin-bottom: 20px; font-family: sans-serif; line-height: 1.5;'>"
                f"  <div style='margin-bottom: 8px; display:flex; justify-content:space-between;'>"
                f"      <span style='color: #00D2FF; font-weight: bold; font-size: 12px; text-transform: uppercase; letter-spacing: 1px;'>PDF LOADED</span>"
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
            
            segments = [s.strip().replace('\n', ' ') for s in re.split(r"\n(?=\[?\d+\]?[\.\s]|\b[A-Z][a-z]+, [A-Z]\.)", content) if len(s.strip()) > 15]
            
            total = len(segments)
            self.initial_count.emit(total)
            history = {} 
            base_item_style = "padding: 12px 0; border-bottom: 1px solid #222; font-family: sans-serif; font-size: 12px; line-height: 1.4;"

            for i, seg in enumerate(segments):
                idx = i + 1
                self.progress_update.emit(idx, total)
                
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
                        
                        # --- MODIFICAÇÃO PARA PUXAR O ABSTRACT COM FALLBACK ---
                        abstract_raw = item.get('abstract', '')
                        
                        # Se o Crossref não tiver o abstract, tenta o Semantic Scholar e OpenAlex
                        if not abstract_raw and res_doi:
                            abstract_raw = self.get_abstract_fallback(res_doi)
                            
                        abstract_clean = self.clean_abstract(abstract_raw)
                        # ------------------------------------------------------

                        is_duplicate = uid in history
                        
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

class DonutChart(QWidget):
    def __init__(self, title):
        super().__init__(); self.title, self.data = title, []
        self.setMinimumHeight(320); self.setMouseTracking(True)
        self.colors = [QColor("#00D2FF"), QColor("#7B2FF7"), QColor("#00FFAA"), QColor("#FFCC00"), QColor("#FF3366"), QColor("#FFFFFF")]
    def update_data(self, d): self.data = sorted(d.items(), key=lambda x: x[1], reverse=True)[:6]; self.update()
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
                    QToolTip.setFont(QFont("Segoe UI", 11)); QToolTip.showText(event.globalPos(), f"<b>{label}</b>: {value}"); return
                start += span
        QToolTip.hideText()
    def paintEvent(self, event):
        p = QPainter(self); p.setRenderHint(QPainter.Antialiasing)
        p.setPen(QColor("#00D2FF")); p.setFont(QFont("Segoe UI", 12, QFont.Bold))
        p.drawText(0, 0, self.width(), 30, Qt.AlignCenter, self.title.upper())
        if not self.data: return
        size = 220; rect = QRectF(self.width()//2 - size//2, self.height()//2 - size//2 + 10, size, size)
        total, start = sum(v for k,v in self.data), 90*16
        for i, (l, v) in enumerate(self.data):
            span = -int((v/total)*360*16)
            p.setBrush(self.colors[i%len(self.colors)]); p.setPen(Qt.NoPen); p.drawPie(rect, start, span); start += span
        p.setBrush(QColor("#050505")); p.drawEllipse(self.width()//2 - 120//2, self.height()//2 - 120//2 + 10, 120, 120)
        p.setPen(QColor("#FFF")); p.setFont(QFont("Segoe UI", 18, QFont.Bold)); p.drawText(rect, Qt.AlignCenter, str(total))

class BarChart(QWidget):
    def __init__(self, title, color_hex="#00D2FF"):
        super().__init__(); self.title, self.data, self.rects = title, [], []
        self.bar_color = QColor(color_hex); self.setMinimumHeight(300); self.setMouseTracking(True); self.margin_left = 200
    def update_data(self, d): self.data = d.most_common(8); self.update()
    def mouseMoveEvent(self, event):
        for rect, label, value in self.rects:
            if rect.contains(event.pos()):
                QToolTip.setFont(QFont("Segoe UI", 11)); QToolTip.showText(event.globalPos(), f"<b>{label}</b>: {value}"); return
        QToolTip.hideText()
    def paintEvent(self, event):
        p = QPainter(self); p.setRenderHint(QPainter.Antialiasing)
        p.setPen(QColor("#00D2FF")); p.setFont(QFont("Segoe UI", 12, QFont.Bold))
        p.drawText(self.margin_left, 0, self.width()-self.margin_left, 30, Qt.AlignLeft, self.title.upper())
        if not self.data: return
        max_v = self.data[0][1]; start_y, bar_h, gap = 40, 20, 10; self.rects = []
        for l, v in self.data:
            w = int((v/max_v)*(self.width() - self.margin_left - 50))
            r = QRectF(self.margin_left, start_y, w, bar_h); self.rects.append((r, l, v))
            p.setBrush(self.bar_color); p.setPen(Qt.NoPen); p.drawRect(r)
            p.setPen(QColor("#AAA")); p.setFont(QFont("Segoe UI", 9))
            label_text = str(l); label_text = label_text[:28] + "..." if len(label_text) > 30 else label_text
            p.drawText(5, int(start_y + bar_h/2 + 4), self.margin_left - 10, bar_h, Qt.AlignRight, label_text)
            p.setPen(QColor("#FFF")); p.setFont(QFont("Segoe UI", 9, QFont.Bold))
            p.drawText(int(self.margin_left + w + 10), int(start_y + bar_h - 4), str(v))
            start_y += (bar_h + gap)


class EditableDetailDialog(QDialog):
    def __init__(self, data, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Detalhes: {data.get('titulo', 'N/A')[:50]}...")
        self.resize(700, 600)
        self.data = data
        self.inputs = {} # Vai guardar as referências dos campos que ficarem editáveis

        # Estilo base da janela replicando o seu visual
        self.setStyleSheet("""
            QDialog { background-color: #050505; color: #FFF; font-family: 'Segoe UI', sans-serif; }
            QLabel { border: none; background: transparent; }
            /* Estilo das caixas editáveis (só aparecem pros N/A) */
            QLineEdit, QTextEdit { 
                background: #1A1A1A; border: 1px dashed #FF3366; 
                border-radius: 4px; color: #FFF; padding: 6px;
            }
            QLineEdit:focus, QTextEdit:focus { border: 1px solid #00D2FF; background: #222; }
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(20)

        # ==========================================
        # BLOCO SUPERIOR: REFERENCE OVERVIEW (Fundo #111, Borda Azul)
        # ==========================================
        overview_frame = QFrame()
        overview_frame.setStyleSheet("""
            QFrame { background-color: #111; border-radius: 5px; border-left: 5px solid #00D2FF; }
        """)
        overview_layout = QVBoxLayout(overview_frame)
        overview_layout.setContentsMargins(15, 15, 15, 15)
        overview_layout.setSpacing(8)

        lbl_tag = QLabel("REFERENCE OVERVIEW")
        lbl_tag.setStyleSheet("color: #00D2FF; font-size: 11px; text-transform: uppercase; font-weight: bold; letter-spacing: 1px;")
        overview_layout.addWidget(lbl_tag)

        # Adiciona os campos (Se for "N/A" vira input, se não, vira texto fixo)
        self.add_field(overview_layout, "titulo", "Título", "font-size: 18px; font-weight: bold; color: #FFF;", QLineEdit)
        self.add_field(overview_layout, "autores", "Autores", "font-size: 14px; color: #DDD;", QLineEdit)

        # Linha Horizontal para Jornal, Ano e DOI
        meta_layout = QHBoxLayout()
        meta_layout.setSpacing(15)
        self.add_field(meta_layout, "jornal", "Journal", "font-size: 13px; color: #AAA;", QLineEdit)
        self.add_field(meta_layout, "ano", "Year", "font-size: 13px; color: #AAA;", QLineEdit)
        self.add_field(meta_layout, "doi", "DOI", "font-size: 13px; color: #00D2FF;", QLineEdit)
        meta_layout.addStretch() # Empurra os itens pra esquerda
        overview_layout.addLayout(meta_layout)

        layout.addWidget(overview_frame)

        # ==========================================
        # BLOCO INFERIOR: ABSTRACT
        # ==========================================
        abstract_layout = QVBoxLayout()
        abstract_layout.setSpacing(10)
        
        lbl_abs = QLabel("ABSTRACT")
        lbl_abs.setStyleSheet("color: #00D2FF; font-size: 12px; font-weight: bold; text-transform: uppercase; letter-spacing: 1px;")
        abstract_layout.addWidget(lbl_abs)

        self.add_field(abstract_layout, "abstract", "Abstract", "font-size: 13px; color: #CCC; line-height: 1.5;", QTextEdit, is_large=True)
        layout.addLayout(abstract_layout)

        # ==========================================
        # BOTÕES: SALVAR / CANCELAR
        # ==========================================
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        
        btn_cancel = QPushButton("Fechar")
        btn_cancel.setStyleSheet("background: #222; color: #FFF; font-weight: bold; padding: 8px 20px; border-radius: 4px; border: none;")
        btn_cancel.clicked.connect(self.reject)
        
        self.btn_save = QPushButton("Salvar Alterações")
        self.btn_save.setStyleSheet("background: #00D2FF; color: #000; font-weight: bold; padding: 8px 20px; border-radius: 4px; border: none;")
        self.btn_save.clicked.connect(self.accept)
        
        # Só mostra o botão salvar se houver campos editáveis
        if not self.inputs:
            self.btn_save.hide()

        btn_layout.addWidget(btn_cancel)
        btn_layout.addWidget(self.btn_save)
        layout.addLayout(btn_layout)

    def add_field(self, parent_layout, key, placeholder, label_style, widget_type, is_large=False):
        val = self.data.get(key, "N/A")
        
        # SE ESTIVER EM BRANCO OU N/A -> MODO DE EDIÇÃO
        if val == "N/A" or not val.strip():
            widget = widget_type()
            widget.setPlaceholderText(f"Inserir {placeholder} (Não encontrado no PDF)...")
            
            if widget_type == QLineEdit:
                widget.setStyleSheet(f"font-size: 13px; {label_style}")
            else:
                widget.setStyleSheet("font-size: 13px; padding: 10px; color: #CCC;")
                widget.setMinimumHeight(180)
            
            self.inputs[key] = widget
            parent_layout.addWidget(widget)
            
        # SE TIVER INFORMAÇÃO -> MODO APENAS LEITURA (TEXTO PURO)
        else:
            widget = QLabel(val)
            widget.setStyleSheet(label_style)
            if is_large:
                widget.setWordWrap(True)
                widget.setAlignment(Qt.AlignTop | Qt.AlignLeft)
            parent_layout.addWidget(widget)

    def get_updated_data(self):
        # Mescla os dados antigos com o que foi digitado nas caixas
        updated = self.data.copy()
        for key, widget in self.inputs.items():
            if isinstance(widget, QLineEdit):
                new_val = widget.text().strip()
            else:
                new_val = widget.toPlainText().strip()
            
            updated[key] = new_val if new_val else "N/A"
        return updated


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
            
            # --- ESTILO GLOBAL ---
            self.setStyleSheet("""
                QMainWindow, QWidget { background: #050505; color: #FFF; font-size: 14px; font-family: 'Segoe UI', Arial; }
                QTableWidget { background: #020202; gridline-color: #111; color: #CCC; border: none; font-size: 12px; }
                QHeaderView::section { background: #0A0A0A; color: #00D2FF; font-weight: bold; padding: 8px; border: none; font-size: 12px; }
                
                QPushButton { background: #111; border: 1px solid #333; color: #BBB; padding: 8px 15px; font-weight: bold; font-size: 13px; min-width: 100px; border-radius: 6px; }
                QPushButton:hover { background: #222; color: #FFF; border-color: #555; }
                
                /* ESTILO DOS BOTÕES DE ABA (CANTO SUPERIOR) */
                QPushButton#CornerTab { 
                    background: #0A0A0A; 
                    color: #888; 
                    border: 1px solid #1A1A1A; 
                    border-bottom: none; 
                    border-top-left-radius: 6px; 
                    border-top-right-radius: 6px; 
                    border-bottom-left-radius: 0px; 
                    border-bottom-right-radius: 0px;
                    padding: 12px 20px; 
                    font-size: 11px;
                    letter-spacing: 1px;
                    text-transform: uppercase;
                    margin-top: 5px; 
                    min-width: 130px; 
                }
                QPushButton#CornerTab:hover { 
                    background: #151515; 
                    color: #00D2FF; 
                    border-color: #333; 
                }

                QPushButton#ActionBtn { background: #003344; color: #00D2FF; border: 1px solid #005577; }
                QPushButton#ActionBtn:hover { background: #004455; }

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
            
            self.btn_load = QPushButton("+ IMPORT PDF")
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

            self.btn_sync_table = QPushButton("Refresh")
            self.btn_sync_table.setObjectName("CornerTab")
            self.btn_sync_table.clicked.connect(self.update_table_view)

            self.btn_imp_bib = QPushButton("IMPORT .bib")
            self.btn_imp_bib.setObjectName("CornerTab")
            self.btn_imp_bib.clicked.connect(self.import_bibtex_file)

            self.btn_bib = QPushButton("EXPORT .bib")
            self.btn_bib.setObjectName("CornerTab")
            self.btn_bib.clicked.connect(self.export_bibtex)

            self.corner_db_layout.addWidget(self.btn_sync_table)
            self.corner_db_layout.addWidget(self.btn_imp_bib)
            self.corner_db_layout.addWidget(self.btn_bib)
            self.corner_db_widget.hide() 

            # -- GRUPO B: DASHBOARD BUTTONS --
            self.corner_dash_widget = QWidget()
            self.corner_dash_layout = QHBoxLayout(self.corner_dash_widget)
            self.corner_dash_layout.setContentsMargins(0, 0, 10, 0)
            self.corner_dash_layout.setSpacing(5)

            self.btn_sync_dash = QPushButton("Refresh")
            self.btn_sync_dash.setObjectName("CornerTab")
            self.btn_sync_dash.clicked.connect(self.update_dashboard_view)

            self.btn_xlsx = QPushButton("EXPORT .xslx")
            self.btn_xlsx.setObjectName("CornerTab")
            self.btn_xlsx.clicked.connect(self.export_excel)

            self.corner_dash_layout.addWidget(self.btn_sync_dash)
            self.corner_dash_layout.addWidget(self.btn_xlsx)
            self.corner_dash_widget.hide()

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

            # Estilo de fonte compacta para a tabela
            table_font = QFont("Segoe UI", 9)
            self.table.setFont(table_font)
            self.table.horizontalHeader().setFont(table_font)

            self.table.setTextElideMode(Qt.ElideRight)
            self.table.setSelectionBehavior(QTableWidget.SelectRows)

            # Altura das linhas compacta
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

            # ADICIONAR ABAS
            self.tabs.addTab(self.tab1_widget, "TERMINAL") # Index 0
            self.tabs.addTab(self.tab2_widget, "DATABASE")    # Index 1
            self.tabs.addTab(self.tab3_widget, "DASHBOARD")   # Index 2

            # CONECTAR SINAL DE TROCA DE ABA
            self.tabs.currentChanged.connect(self.on_tab_changed)

            # --- RODAPÉ DE PROGRESSO ---
            self.progress_label = QLabel("")
            self.progress_label.setStyleSheet("""
                QLabel {
                    color: #00D2FF; 
                    background: #000000;
                    padding: 0px; 
                    font-family: 'Consolas', 'Monaco', monospace; 
                    font-size: 14px;
                    font-weight: bold;
                    border-top: 1px solid #1A1A1A;
                }
            """)
            self.progress_label.setFixedHeight(50) 
            self.progress_label.setAlignment(Qt.AlignCenter)
            self.progress_label.hide()

            main_layout.addWidget(self.tabs)
            main_layout.addWidget(self.progress_label)
            self.setCentralWidget(main_container)

    def on_tab_changed(self, index):
            # Esconde ambos primeiro para evitar bugs visuais
            self.corner_db_widget.hide()
            self.corner_dash_widget.hide()

            if index == 1:
                self.tabs.setCornerWidget(self.corner_db_widget, Qt.TopRightCorner)
                self.corner_db_widget.show()
            elif index == 2:
                self.tabs.setCornerWidget(self.corner_dash_widget, Qt.TopRightCorner)
                self.corner_dash_widget.show()
                self.update_dashboard_view() # Atualiza os gráficos ao entrar na aba
            else:
                empty = QWidget()
                self.tabs.setCornerWidget(empty, Qt.TopRightCorner)

    def setup_dash_widgets(self):
        self.l_dash.setContentsMargins(40, 40, 40, 40); self.l_dash.setSpacing(40)
        h_kpi = QHBoxLayout(); h_kpi.setSpacing(20)
        self.k1 = self.create_kpi("TOTAL REFERÊNCIAS")
        self.k2 = self.create_kpi("ABSTRACTS OBTIDOS") 
        self.k3 = self.create_kpi("ANO MÁX")
        self.k4 = self.create_kpi("FONTES (PDFs)")
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
            # Coleta os dados da linha inteira independentemente de qual célula foi clicada
            try:
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
                print(f"Erro ao capturar dados da célula: {e}")

    def show_abstract_popup(self, data):
        # Criamos um Diálogo customizado em vez de QMessageBox para ter controle total
        dialog = QDialog(self)
        dialog.setWindowTitle(f"Detalhes: {data['titulo'][:50]}...")
        dialog.setMinimumSize(900, 650)
        dialog.setStyleSheet("background-color: #050505;")
        
        layout = QVBoxLayout(dialog)

        # 1. CABEÇALHO (O seu HTML original - Apenas leitura)
        header_browser = QTextBrowser()
        header_browser.setOpenExternalLinks(True)
        header_browser.setFrameStyle(QFrame.NoFrame)
        header_browser.setFixedHeight(200) # Espaço fixo para o topo
        
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
        # 2. CAMPO DE EDIÇÃO DO ABSTRACT
        # Usamos QTextEdit para permitir digitação
        abstract_edit = QTextEdit()
        abstract_text = data['abstract'] if data['abstract'] != 'N/A' else ""
        abstract_edit.setPlainText(abstract_text)
        abstract_edit.setPlaceholderText("Insira abstract here...")
        abstract_edit.setStyleSheet("""
            QTextEdit {
                background-color: #0A0A0A;
                color: #DDD;
                border: 1px solid #222;
                border-radius: 5px;
                padding: 15px;
                
                font-family: 'Segoe UI';
                font-size: 14px;
                line-height: 1.8;
            }
            QTextEdit:focus { border: 1px solid #00D2FF; }
        """)
        layout.addWidget(abstract_edit)

        # 3. BOTÕES DE AÇÃO
        btn_layout = QHBoxLayout()
        
        # Botão de Salvar Alterações na Tabela
        btn_save = QPushButton(" Save to Database")
        btn_save.setCursor(Qt.PointingHandCursor)
        btn_save.setStyleSheet("""
            QPushButton { 
                background: #004433; color: #00FFCC; padding: 10px; 
                border-radius: 5px; border: 1px solid #006644; font-weight: bold; 
            }
            QPushButton:hover { background: #005544; }
        """)
        
        # Botão de Copiar BibTeX (Seu original)
        btn_copy_bib = QPushButton(" Copy BibTeXto dashboard")
        btn_copy_bib.setCursor(Qt.PointingHandCursor)
        btn_copy_bib.setStyleSheet("""
            QPushButton { 
                background: #003344; color: #00D2FF; padding: 10px; 
                border-radius: 5px; border: 1px solid #005577; font-weight: bold; 
            }
            QPushButton:hover { background: #004455; }
        """)

        btn_layout.addWidget(btn_save)
        btn_layout.addWidget(btn_copy_bib)
        layout.addLayout(btn_layout)

        # FUNÇÕES DOS BOTÕES
        def save_changes():
            new_abstract = abstract_edit.toPlainText().strip()
            # Atualiza a memória (live_data)
            row_idx = self.table.currentRow()
            if row_idx >= 0:
                # Atualiza a tabela visualmente (Coluna 8 é o Abstract)
                self.table.setItem(row_idx, 8, QTableWidgetItem(new_abstract))
                # Atualiza o dicionário data para o BibTeX usar o texto novo
                data['abstract'] = new_abstract
            btn_save.setText("✓ Saved!")
            btn_save.setEnabled(False)

        def copy_and_close():
            current_abs = abstract_edit.toPlainText().strip()
            author_key = data['autores'].split(',')[0].strip().split(' ')[0]
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

    def open_pdf(self):
        p, _ = QFileDialog.getOpenFileName(self, "Adicionar PDF ao Banco de Dados", "", "*.pdf")
        if p:
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

            # Spinner Braille de 8 frames (mais fluido)
            spinner_chars = ["⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏"]
            spinner = spinner_chars[self.spin_idx % len(spinner_chars)]
            
            current, total = self.current_progress
            
            # bar_length configurado para ocupar a linha inteira (ajuste conforme sua UI)
            bar_length = 135 
            
            if total == 0:
                # Efeito de "Scanner" suave para inicialização
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
                
                # Caractere de ponta suave (Half-block transition)
                smooth_tip = [" ", "╸", "━"] 
                remainder = int((filled_exact - filled_full) * (len(smooth_tip) - 1))
                tip = smooth_tip[remainder] if filled_full < bar_length else ""
                
                bar_main = "━" * filled_full
                bar_empty = " " * (bar_length - filled_full - (1 if tip else 0))
                
                if current >= total:
                    # Layout de conclusão 100%
                    full_bar = "━" * bar_length
                    msg = f"✔  SYSTEM READY  {full_bar}"
                else:
            # 1. Descobre quantos dígitos o total tem (ex: 50 -> 2, 1500 -> 4)
                    padding = len(str(total))
                    
                    # 2. Formata dinamicamente: o '0' indica preenchimento com zeros, 
                    # e o 'padding' define a largura.
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
        msg = "<br><b style='color:#00FFAA;'>FINISH.</b><br>" if success else "<br><b style='color:#FF3366;'>ERRO.</b>"
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

    def update_dashboard_view(self):
        if not self.live_data: return
        yrs, auths, journals, pubs, files = [], [], [], [], []
        abs_count = 0
        for row in self.live_data:
            files.append(row[0])
            if str(row[3]).isdigit(): yrs.append(str(row[3]))
            row_auths = [a.strip() for a in row[1].split(';') if len(a) > 3]
            auths.extend(row_auths)
            if row[8] and row[8] != "N/A": abs_count += 1
            if row[5] and row[5] != "N/A": journals.append(row[5])
            if row[6] and row[6] != "N/A": pubs.append(row[6])
        self.k1.v.setText(str(len(self.live_data)))
        self.k2.v.setText(str(abs_count))
        self.k4.v.setText(str(len(set(files))))
        if yrs: self.k3.v.setText(str(max(yrs)))
        self.c_yrs.update_data(Counter(yrs))
        self.c_tps.update_data(Counter([row[7] for row in self.live_data]))
        self.c_pub.update_data(Counter(pubs))
        self.b_journal.update_data(Counter(journals))
        self.b_aut.update_data(Counter(auths))


    def export_bibtex(self):
        path, _ = QFileDialog.getSaveFileName(self, "Exportar BibTeX Consolidado", "", "*.bib")
        if path and self.live_data:
            with open(path, "w", encoding="utf-8") as f:
                for r in self.live_data:
                    k = re.sub(r'\W+', '', r[1].split(',')[0]) + str(r[3])
                    abs_field = f"  abstract={{{r[8]}}},\n" if r[8] != "N/A" else ""
                    f.write(f"@article{{{k},\n  author={{{r[1]}}},\n  title={{{r[2]}}},\n  journal={{{r[5]}}},\n  year={{{r[3]}}},\n  doi={{{r[4]}}},\n{abs_field}  publisher={{{r[6]}}},\n  note={{Source File: {r[0]}}}\n}}\n\n")
            QMessageBox.information(self, "Exportar", "Arquivo BibTeX salvo.")

    def export_excel(self):
        path, _ = QFileDialog.getSaveFileName(self, "Exportar Excel Consolidado", "", "*.xlsx")
        if path and self.live_data:
            df = pd.DataFrame(self.live_data, columns=["Arquivo Fonte", "Autores", "Título", "Ano", "DOI", "Jornal", "Publisher", "Tipo", "Abstract"])
            df.to_excel(path, index=False)
            QMessageBox.information(self, "Exportar", "Arquivo Excel salvo.")

    def import_bibtex_file(self):
            """Lê arquivo .bib e utiliza o sistema de animação progress_label."""
            fname, _ = QFileDialog.getOpenFileName(self, "Importar BibTeX", "", "BibTeX Files (*.bib)")
            if not fname: return

            try:
                with open(fname, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                raw_entries = re.findall(r'@(\w+)\s*\{([^@]+)\}', content, re.DOTALL)
                if not raw_entries: return

                # --- ATIVA O SISTEMA DE CARREGAMENTO ---
                self.is_processing = True
                total = len(raw_entries)
                new_entries = []

                for i, (type_str, entry_body) in enumerate(raw_entries):
                    # Atualiza os dados que o seu update_progress_label consome
                    self.current_progress = (i + 1, total)
                    
                    # Campos padrão
                    fields = {'author':'N/A', 'title':'Sem Título', 'year':'n.d.', 'doi':'', 
                            'journal':'N/A', 'publisher':'N/A', 'abstract':''}
                    
                    matches = re.findall(r'(\w+)\s*=\s*(?:\{([^}]*)\}|"([^"]*)"|(\d+))', entry_body, re.IGNORECASE)
                    for k, vb, vq, vr in matches:
                        val = vb or vq or vr
                        val = re.sub(r'\s+', ' ', val.replace('\n', ' ').replace('{','').replace('}','').replace('\\','')).strip()
                        k = k.lower()
                        if k in fields: fields[k] = val
                        elif k == 'booktitle': fields['journal'] = val
                    
                    # Busca de Abstract (Enriquecimento)
                    doi_val = fields['doi'].strip().lower()
                    if len(fields['abstract']) < 20 and doi_val:
                        # Chamamos a API - a UI continuará animando se o timer estiver ativo
                        api_abs = self.fetch_abstract_online(doi_val)
                        if api_abs: fields['abstract'] = api_abs

                    doi_link = f"https://doi.org/{doi_val}" if doi_val else ""
                    
                    row = [
                        os.path.basename(fname), 
                        fields['author'], fields['title'], fields['year'], 
                        doi_link, fields['journal'], fields['publisher'], 
                        type_str.title(), fields['abstract']
                    ]
                    new_entries.append(row)
                    
                    # Força a atualização visual da barra (chama seu def)
                    self.update_progress_label() 
                    QApplication.processEvents()

                # --- FINALIZAÇÃO ---
                self.live_data.extend(new_entries)
                self.update_table_view()
                
                # Deixa a barra em 100% por um momento antes de sumir
                self.current_progress = (total, total)
                self.update_progress_label()
                
                QTimer.singleShot(2000, self.stop_processing_ui) # Função auxiliar para resetar

            except Exception as e:
                self.is_processing = False
                self.update_progress_label()
                QMessageBox.critical(self, "Erro", f"Erro na importação: {e}")

    def stop_processing_ui(self):
            """Reseta os flags de animação."""
            self.is_processing = False
            self.current_progress = (0, 0)
            self.update_progress_label()

    def fetch_abstract_online(self, doi):
            """Busca abstract em múltiplas fontes para preencher lacunas do BibTeX."""
            if not doi: return ""
            # 1. Crossref (Lida bem com metadados brutos de editoras)
            try:
                res = requests.get(f"https://api.crossref.org/works/{doi}", timeout=3)
                if res.status_code == 200:
                    abs_raw = res.json().get('message', {}).get('abstract', '')
                    if abs_raw:
                        clean = re.sub(r'<[^>]+>', '', abs_raw)
                        return re.sub(r'\s+', ' ', clean).strip()
            except: pass

            # 2. Semantic Scholar (Excelente para Springer)
            try:
                res = requests.get(f"https://api.semanticscholar.org/graph/v1/paper/DOI:{doi}?fields=abstract", timeout=3)
                if res.status_code == 200:
                    return res.json().get('abstract', '')
            except: pass

            # 3. OpenAlex (Reconstrução de índice invertido)
            try:
                res = requests.get(f"https://api.openalex.org/works/https://doi.org/{doi}", timeout=3)
                if res.status_code == 200:
                    inv = res.json().get('abstract_inverted_index')
                    if inv:
                        word_idx = []
                        for w, pos in inv.items():
                            for p in pos: word_idx.append((p, w))
                        word_idx.sort()
                        return " ".join([x[1] for x in word_idx])
            except: pass
            return ""

if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = BiblioApp()
    w.show()
    sys.exit(app.exec_())