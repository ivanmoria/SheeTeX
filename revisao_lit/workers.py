import re, math, html, os, requests
from bs4 import BeautifulSoup
from PyQt5.QtWidgets import ( QWidget, QVBoxLayout, QPushButton, QLabel, QHBoxLayout, QFrame, QToolTip,QDialog, QLineEdit, QTextEdit )
from PyQt5.QtCore import QThread, pyqtSignal, Qt, QRectF, QPoint, QRect
from PyQt5.QtGui import QPainter, QColor, QFont
from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextContainer, LAParams
from habanero import Crossref
import os
import re
import html
import requests
from bs4 import BeautifulSoup
from PyQt5.QtCore import QThread, pyqtSignal


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

    def resolve_source_metadata(self, journal_raw, publisher_raw):
        """
        Unifica Journal e Publisher para evitar N/A se um deles existir.
        """
        j = self.clean_text(journal_raw)
        p = self.clean_text(publisher_raw)
        
        # Lista de valores considerados vazios
        invalid = ["", "n/a", "na", "unknown"]

        j_invalid = j.lower() in invalid
        p_invalid = p.lower() in invalid

        if j_invalid and not p_invalid:
            j = p  # Journal assume o valor do Publisher
        elif p_invalid and not j_invalid:
            p = j  # Publisher assume o valor do Journal
        elif j_invalid and p_invalid:
            j = "N/A"
            p = "N/A"
            
        return j, p

    def scrape_page_for_abstract(self, doi):
        # ... (código inalterado para brevidade, manter sua lógica original de scraping)
        url = f"https://doi.org/{doi}"
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8'
        }
        try:
            res = requests.get(url, headers=headers, timeout=10, allow_redirects=True)
            if res.status_code != 200: return ""
            soup = BeautifulSoup(res.text, 'html.parser')

            meta_tags = [
                soup.find('meta', attrs={'name': 'citation_abstract'}),
                soup.find('meta', attrs={'name': 'dc.description'}),
                soup.find('meta', attrs={'name': 'description'}),
                soup.find('meta', property='og:description')
            ]
            for tag in meta_tags:
                if tag and tag.get('content'):
                    text = tag.get('content').strip()
                    if len(text) > 100: return text

            target_elements = soup.find_all(['div', 'section', 'p'], class_=re.compile(r'abstract|resumo', re.I))
            target_elements += soup.find_all(['div', 'section', 'p'], id=re.compile(r'abstract|resumo', re.I))
            for el in target_elements:
                text = el.get_text(separator=' ', strip=True)
                if len(text) > 150: 
                    text = re.sub(r'^(Abstract|Resumo)\s*[:\-]?\s*', '', text, flags=re.IGNORECASE)
                    return text
            
            # Fallback headers logic...
            headers_tags = soup.find_all(['h2', 'h3', 'strong', 'h1', 'span'], string=re.compile(r'(?i)^\s*(Abstract|Resumo)\s*$'))
            for header in headers_tags:
                nxt = header.find_next_sibling(['p', 'div'])
                if nxt:
                    text = nxt.get_text(separator=' ', strip=True)
                    if len(text) > 100: return text
        except Exception: pass
        return ""

    def get_abstract_fallback(self, doi):
        if not doi: return ""
        try:
            s2_url = f"https://api.semanticscholar.org/graph/v1/paper/DOI:{doi}?fields=abstract"
            s2_res = requests.get(s2_url, timeout=5) 
            if s2_res.status_code == 200:
                data = s2_res.json()
                if data.get('abstract'): return data['abstract']
        except Exception: pass 
        
        try:
            oa_url = f"https://api.openalex.org/works/https://doi.org/{doi}"
            headers = {'User-Agent': 'mailto:pesquisador@universidade.edu'}
            oa_res = requests.get(oa_url, headers=headers, timeout=5)
            if oa_res.status_code == 200:
                data = oa_res.json()
                abstract_inverted = data.get('abstract_inverted_index')
                if abstract_inverted:
                    return self.reconstruct_openalex_abstract(abstract_inverted)
        except Exception: pass

        return self.scrape_page_for_abstract(doi)

    def reconstruct_openalex_abstract(self, inverted_index):
        if not inverted_index: return ""
        word_index = []
        for word, positions in inverted_index.items():
            for pos in positions: word_index.append((pos, word))
        word_index.sort(key=lambda x: x[0])
        return " ".join([word for pos, word in word_index])

    def run(self):
        # Assumindo Crossref e bibliotecas importadas
        from habanero import Crossref # Exemplo
        from pdfminer.layout import LAParams, LTTextContainer
        from pdfminer.high_level import extract_pages

        cr = Crossref()
        try:
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
                if page_count == 0: first_page_text = page_text
                full_text += page_text
                page_count += 1

            doc_title = os.path.basename(self.pdf_path)
            doc_author = "Autor Desconhecido"
            doc_year = "n.d."
            
            # Tratamento inicial do Documento Principal
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

            doc_journal = "Periódico Não Identificado"
            doc_full_ref = "Referência completa não disponível."

            if main_item:
                doc_title = main_item.get("title", [doc_title])[0]
                auths = main_item.get('author', [])
                if auths: doc_author = "; ".join([f"{a.get('family', '')}, {a.get('given', '')}" for a in auths])
                
                date_parts = main_item.get("published-print", {}).get("date-parts") or \
                            main_item.get("published-online", {}).get("date-parts") or \
                            main_item.get("created", {}).get("date-parts")
                if date_parts: doc_year = str(date_parts[0][0])
                
                # --- MELHORIA AQUI PARA O DOC PRINCIPAL ---
                raw_j = main_item.get("container-title", [""])[0]
                raw_p = main_item.get("publisher", "")
                doc_journal, _ = self.resolve_source_metadata(raw_j, raw_p)
                
                doi_val = main_item.get("DOI", "")
                doc_doi_link = f"https://doi.org/{doi_val}" if doi_val else ""
                
                doc_full_ref = f"<i>{doc_journal}</i>  • {doc_year}. "
                if doc_doi_link: doc_full_ref += f"  •  DOI: <a href='{doc_doi_link}' style='color:#00D2FF; text-decoration:none;'>{doi_val}</a>"

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
                        
                        # --- MODIFICAÇÃO PRINCIPAL AQUI ---
                        raw_container = item.get("container-title", [""])
                        raw_j_item = raw_container[0] if raw_container else ""
                        raw_p_item = item.get("publisher", "")
                        
                        journal, pub = self.resolve_source_metadata(raw_j_item, raw_p_item)
                        # ----------------------------------

                        res_doi = item.get('DOI', '').lower()
                        link = f"https://doi.org/{res_doi}" if res_doi else ""
                        uid = res_doi if res_doi else f"{auth_str}{title}{year}".lower()
                        
                        abstract_raw = item.get('abstract', '')
                        if not abstract_raw and res_doi:
                            abstract_raw = self.get_abstract_fallback(res_doi)
                            
                        abstract_clean = self.clean_abstract(abstract_raw)
                        is_duplicate = uid in history
                        
                        has_abstract = len(abstract_clean) > 20 and abstract_clean != "N/A"
                        abstract_status = "<span style='color:#00D2FF; font-weight:bold;'>Found</span>" if has_abstract else "<span style='color:#FF3366; font-weight:bold;'>Not Found</span>"

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
                            f"  <div style='color:#777; margin-top: 2px;'>{journal} • {year} • DOI: {doi_html} • Abstract: {abstract_status}</div>"
                            f"</div>"
                        )

                        if not is_duplicate:
                            history[uid] = title
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
# ENGINE: PROCESSAMENTO DE BIBTEX (WORKER)
# =============================================================================
class BibtexWorker(QThread):
    initial_count = pyqtSignal(int)
    progress_update = pyqtSignal(int, int)
    item_completed = pyqtSignal(list)
    log_result = pyqtSignal(str)
    header_ready = pyqtSignal(str)
    finished_signal = pyqtSignal(int)

    def __init__(self, fname):
        super().__init__()
        self.fname = fname
        self.filename_base = os.path.basename(fname)

    def clean_doi(self, doi_str):
        if not doi_str: return ""
        doi = doi_str.replace("https://doi.org/", "").replace("http://doi.org/", "").strip()
        doi = doi.replace("{", "").replace("}", "").strip()
        return doi

    def clean_abstract(self, text):
        if not text: return "N/A"
        text = html.unescape(text)
        clean = re.sub(r'<[^>]+>', '', text)
        return re.sub(r'\s+', ' ', clean).strip()

    def scrape_page_for_abstract(self, doi):
        # ... Mesma lógica de scrape ...
        url = f"https://doi.org/{doi}"
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8'
        }
        try:
            res = requests.get(url, headers=headers, timeout=10, allow_redirects=True)
            if res.status_code != 200: return ""
            soup = BeautifulSoup(res.text, 'html.parser')

            meta_tags = [
                soup.find('meta', attrs={'name': 'citation_abstract'}),
                soup.find('meta', attrs={'name': 'dc.description'}),
                soup.find('meta', attrs={'name': 'description'})
            ]
            for tag in meta_tags:
                if tag and tag.get('content'):
                    text = tag.get('content').strip()
                    if len(text) > 100: return text
            
            # (Simplificando repetição: mantenha a lógica completa do seu código original aqui se quiser, ou a mesma do DeepExtractionWorker)
            return "" 
        except: pass
        return ""

    def get_abstract_fallback(self, doi):
        doi = self.clean_doi(doi)
        if not doi: return ""
        
        try:
            s2_url = f"https://api.semanticscholar.org/graph/v1/paper/DOI:{doi}?fields=abstract"
            s2_res = requests.get(s2_url, timeout=5) 
            if s2_res.status_code == 200:
                data = s2_res.json()
                if data.get('abstract'): return data['abstract']
        except: pass 

        try:
            oa_url = f"https://api.openalex.org/works/https://doi.org/{doi}"
            headers = {'User-Agent': 'mailto:pesquisador@universidade.edu'}
            oa_res = requests.get(oa_url, headers=headers, timeout=5)
            if oa_res.status_code == 200:
                data = oa_res.json()
                inverted = data.get('abstract_inverted_index')
                if inverted:
                    word_index = []
                    for word, positions in inverted.items():
                        for pos in positions: word_index.append((pos, word))
                    word_index.sort(key=lambda x: x[0])
                    return " ".join([word for pos, word in word_index])
        except: pass
        
        return self.scrape_page_for_abstract(doi)

    def run(self):
        try:
            with open(self.fname, 'r', encoding='utf-8') as f:
                content = f.read()
            
            raw_entries = re.findall(r'@(\w+)\s*\{([^@]+)\}', content, re.DOTALL)
            if not raw_entries:
                self.finished_signal.emit(0)
                return

            total = len(raw_entries)
            self.initial_count.emit(total)
            header_html = (
                f"<div style='padding: 15px 0; border-bottom: 2px solid #333; margin-bottom: 20px; font-family: sans-serif;'>"
                f"  <span style='color: #FF9900; font-weight: bold; font-size: 12px; text-transform: uppercase;'>BIBTEX IMPORT</span>"
                f"  <div style='color: #FFF; font-size: 14px; font-weight: bold; margin-top: 5px;'>{self.filename_base}</div>"
                f"</div>"
            )
            self.header_ready.emit(header_html)
            
            base_item_style = "padding: 12px 0; border-bottom: 1px solid #222; font-family: sans-serif; font-size: 12px; line-height: 1.4;"

            for i, (type_str, entry_body) in enumerate(raw_entries):
                # Inicializa com strings vazias em vez de N/A fixo para facilitar a lógica de limpeza
                fields = {'author':'N/A', 'title':'Sem Título', 'year':'n.d.', 'doi':'', 
                        'journal':'', 'publisher':'', 'abstract':''}
                
                matches = re.findall(r'(\w+)\s*=\s*(?:\{([^}]*)\}|"([^"]*)"|(\d+))', entry_body, re.IGNORECASE)
                for k, vb, vq, vr in matches:
                    val = vb or vq or vr
                    val = re.sub(r'\s+', ' ', val.replace('\n', ' ').replace('{','').replace('}','').replace('\\','')).strip()
                    k = k.lower()
                    if k in fields: fields[k] = val
                    elif k == 'booktitle': fields['journal'] = val

                # --- MELHORIA: Lógica de Unificação Journal/Publisher ---
                j = fields['journal']
                p = fields['publisher']
                
                if not j and p: j = p
                elif not p and j: p = j
                elif not j and not p:
                     j = "N/A"
                     p = "N/A"
                
                fields['journal'] = j
                fields['publisher'] = p
                # --------------------------------------------------------

                doi_val = self.clean_doi(fields['doi'])
                
                if (len(fields['abstract']) < 20 or fields['abstract'] == "N/A") and doi_val:
                    online_abs = self.get_abstract_fallback(doi_val)
                    if online_abs:
                        fields['abstract'] = self.clean_abstract(online_abs)
                
                if not fields['abstract']: fields['abstract'] = "N/A"

                doi_link = f"https://doi.org/{doi_val}" if doi_val else ""
                
                row = [
                    self.filename_base, fields['author'], fields['title'], fields['year'], 
                    doi_link, fields['journal'], fields['publisher'], 
                    type_str.title(), fields['abstract']
                ]
                
                self.item_completed.emit(row)
                
                idx = i + 1
                
                has_abstract = len(fields['abstract']) > 20 and fields['abstract'] != "N/A"
                abstract_status = "<span style='color:#00D2FF; font-weight:bold;'>Found</span>" if has_abstract else "<span style='color:#FF3366; font-weight:bold;'>Not Found</span>"
                doi_html = f"<a href='{doi_link}' style='color:#00D2FF; text-decoration:underline;'>{doi_val}</a>" if doi_val else "N/A"

                detail_log = (
                    f"<div style='{base_item_style}'>"
                    f"  <div style='margin-bottom: 4px;'>"
                    f"    <span style='color:#FF9900; font-family:Consolas; font-weight:bold;'>[{idx:03d}/{total:03d}]</span> "
                    f"    <b style='color:#FFF; font-size: 13px;'>{fields['title']}</b>"
                    f"  </div>"
                    f"  <div style='color:#BBB;'>{fields['author']}</div>"
                    f"  <div style='color:#777; margin-top: 2px;'>{fields['journal']} • {fields['year']} • DOI: {doi_html} • Abstract: {abstract_status}</div>"
                    f"</div>"
                )
                
                self.log_result.emit(detail_log)
                self.progress_update.emit(idx, total)

            self.finished_signal.emit(total)

        except Exception as e:
            self.log_result.emit(f"<div style='color:red;'>Erro: {str(e)}</div>")
            self.finished_signal.emit(-1)
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
        super().__init__()
        self.title, self.data, self.rects = title, [], []
        self.bar_color = QColor(color_hex)
        self.setMinimumHeight(350) 
        self.setMouseTracking(True)
        self.margin_left = 300 # Margem padrão aumentada

    def update_data(self, d): 
        self.data = d.most_common(10) # Aumentado para 10 itens
        self.update()

    def mouseMoveEvent(self, event):
        for rect, label, value in self.rects:
            if rect.contains(event.pos()):
                QToolTip.setFont(QFont("Segoe UI", 11))
                # Tooltip sempre ajuda a ler se o gráfico estiver muito pequeno
                QToolTip.showText(event.globalPos(), f"<div style='width:300px'><b>{label}</b>: {value}</div>")
                return
        QToolTip.hideText()

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        
        # Título do Gráfico
        p.setPen(QColor("#00D2FF"))
        p.setFont(QFont("Segoe UI", 11, QFont.Bold))
        p.drawText(self.margin_left, 0, self.width()-self.margin_left, 30, Qt.AlignLeft, self.title.upper())
        
        if not self.data: return
        
        max_v = self.data[0][1]
        start_y, bar_h, gap = 50, 25, 25 # Aumentamos o gap para caber o texto multiline
        self.rects = []
        
        for l, v in self.data:
            # Cálculo da barra
            w = int((v/max_v)*(self.width() - self.margin_left - 60))
            r = QRectF(self.margin_left, start_y, w, bar_h)
            self.rects.append((r, l, v))
            
            # Desenha a Barra
            p.setBrush(self.bar_color)
            p.setPen(Qt.NoPen)
            p.drawRect(r)
            
            # Desenha o Label (NOME INTEIRO COM QUEBRA DE LINHA)
            p.setPen(QColor("#CCC"))
            p.setFont(QFont("Segoe UI", 9))
            
            # Definimos um retângulo para o texto à esquerda
            # Qt.TextWordWrap permite que o nome ocupe mais de uma linha se necessário
            label_rect = QRect(5, int(start_y - 5), self.margin_left - 15, bar_h + 15)
            p.drawText(label_rect, Qt.AlignRight | Qt.AlignVCenter | Qt.TextWordWrap, str(l))
            
            # Desenha o Valor (Número)
            p.setPen(QColor("#FFF"))
            p.setFont(QFont("Segoe UI", 9, QFont.Bold))
            p.drawText(int(self.margin_left + w + 10), int(start_y + bar_h - 6), str(v))
            
            # Incremento para a próxima linha (espaço maior para acomodar quebras de linha)
            start_y += (bar_h + gap)
            
        # Ajusta a altura mínima dinamicamente para não cortar o gráfico
        if self.minimumHeight() < start_y:
            self.setMinimumHeight(start_y + 20)

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


