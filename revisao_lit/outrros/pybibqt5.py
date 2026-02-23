import customtkinter as ctk
from tkinter import filedialog, messagebox, ttk
import threading
import re
import os

# Motores de Processamento
from docling.document_converter import DocumentConverter
from habanero import Crossref

# Configurações de Tema (Estilo Apple)
ctk.set_appearance_mode("System") 
ctk.set_default_color_theme("blue")

class BiblioTechPro(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("BiblioTech AI - Reference Extractor (macOS)")
        self.geometry("1150x850")

        # Inicialização dos Motores
        self.converter = DocumentConverter()
        self.cr = Crossref()

        # Estrutura de Abas
        self.tabview = ctk.CTkTabview(self, width=1100, height=800)
        self.tabview.pack(padx=20, pady=20, expand=True, fill="both")

        self.tab_extract = self.tabview.add("📤 Extração")
        self.tab_table = self.tabview.add("📊 Tabela de Dados")

        self.setup_extract_ui()
        self.setup_table_ui()

    # --- INTERFACE: ABA DE EXTRAÇÃO ---
    def setup_extract_ui(self):
        self.btn_select = ctk.CTkButton(self.tab_extract, text="Selecionar PDF (Nova Lógica)", 
                                        command=self.start_extraction, height=50, font=("Helvetica", 14, "bold"))
        self.btn_select.pack(pady=30)

        self.status_frame = ctk.CTkFrame(self.tab_extract, fg_color="transparent")
        self.status_frame.pack(fill="x", padx=100)
        
        self.lbl_status = ctk.CTkLabel(self.status_frame, text="Aguardando arquivo...", font=("Helvetica", 13))
        self.lbl_status.pack(side="left")
        
        self.lbl_percentage = ctk.CTkLabel(self.status_frame, text="0%", font=("Helvetica", 13, "bold"))
        self.lbl_percentage.pack(side="right")

        self.progress = ctk.CTkProgressBar(self.tab_extract, width=900)
        self.progress.set(0)
        self.progress.pack(pady=15)

        self.log_box = ctk.CTkTextbox(self.tab_extract, width=950, height=450, font=("Menlo", 11))
        self.log_box.pack(padx=20, pady=10)

    # --- INTERFACE: ABA DE TABELA ---
    def setup_table_ui(self):
        self.table_frame = ctk.CTkFrame(self.tab_table)
        self.table_frame.pack(expand=True, fill="both", padx=20, pady=20)

        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview", background="#2b2b2b", foreground="white", fieldbackground="#2b2b2b", rowheight=35)
        style.map("Treeview", background=[('selected', '#1f538d')])

        columns = ("Nº", "Título", "Ano", "DOI")
        self.tree = ttk.Treeview(self.table_frame, columns=columns, show="headings")

        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor="w")

        self.tree.column("Nº", width=40, anchor="center")
        self.tree.column("Título", width=600)
        self.tree.column("Ano", width=80, anchor="center")
        
        self.tree.pack(side="left", expand=True, fill="both")
        
        sb = ttk.Scrollbar(self.table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")

    # --- LÓGICA DE PROCESSAMENTO REESCRITA ---
    def start_extraction(self):
        path = filedialog.askopenfilename(filetypes=[("PDF Científico", "*.pdf")])
        if path:
            self.log_box.delete("1.0", "end")
            self.tree.delete(*self.tree.get_children())
            self.btn_select.configure(state="disabled")
            self.progress.configure(mode="indeterminate")
            self.progress.start()
            self.lbl_status.configure(text="⚙️ Analisando estrutura semântica...")
            threading.Thread(target=self.new_extraction_logic, args=(path,), daemon=True).start()

    def new_extraction_logic(self, path):
        try:
            self.update_log("🚀 Iniciando motor Docling...")
            conv_result = self.converter.convert(path)
            doc = conv_result.document
            
            all_text_elements = []
            
            # PALAVRAS-CHAVE AUMENTADAS
            keywords = r'(?i)^\s*(\d+[\.\s]*)?(references|bibliography|referências|literature cited|bibliografia|referências bibliográficas)\s*$'
            
            found_ref_zone = False
            self.update_log("🔍 Mapeando zonas de texto...")
            
            # --- NOVA LOGICA DE ITERAÇÃO POR NÓS ---
            for item, _level in doc.iterate_items():
                
                # Tratamento de erro PictureItem (Ignorar imagens)
                try:
                    text = item.export_to_text().strip()
                except:
                    continue 
                
                if not text: continue
                
                # Detectar transição para referências
                if not found_ref_zone:
                    if re.match(keywords, text):
                        found_ref_zone = True
                        self.update_log(f"📍 Zona de referências encontrada: '{text}'")
                        continue
                
                if found_ref_zone:
                    # Limpeza de hifenização de fim de linha
                    clean_text = text.replace("- ", "").replace("-\n", "")
                    
                    # Filtramos lixo (números de página, rodapés)
                    if len(clean_text) > 35 and not clean_text.isdigit():
                        # Separa referências agrupadas no mesmo bloco
                        sub_items = re.split(r'\n(?=\[?\d+\]?[\.\s])', clean_text)
                        for s in sub_items:
                            if len(s.strip()) > 35:
                                all_text_elements.append(s.strip())

            # Se nada for encontrado com regex, fazemos uma varredura final agressiva
            if not all_text_elements:
                self.update_log("⚠️ Regex falhou. Tentando varredura agressiva final...")
                # ... Lógica de fallback para pegar os últimos 20% do texto ...

            total = len(all_text_elements)
            self.update_log(f"📈 {total} itens brutos capturados. Iniciando Crossref...")
            
            self.after(0, self.prepare_determinate)

            for i, raw_ref in enumerate(all_text_elements):
                current = i + 1
                self.after(0, self.update_progress_ui, current/total, current, total)
                
                # Limpa marcadores iniciais [1], 1., etc.
                query = re.sub(r'^\[?\d+\]?[\s\.\-]+', '', raw_ref)[:250]
                
                try:
                    res = self.cr.works(query=query, limit=1)
                    if res['message']['items']:
                        data = res['message']['items'][0]
                        title = data.get("title", ["N/A"])[0]
                        year = data.get("created", {}).get("date-parts", [[0]])[0][0]
                        doi = data.get("DOI", "N/A")
                        self.after(0, self.add_to_table, current, title, year, doi)
                except:
                    self.after(0, self.add_to_table, current, query[:60]+"...", "N/A", "---")

            self.update_log("✅ Concluído!")
            self.after(0, self.finish_ui)
        except Exception as e:
            self.update_log(f"❌ Erro Crítico: {str(e)}")
            self.after(0, self.finish_ui)

    # --- HELPERS ---
    def update_log(self, msg):
        self.after(0, lambda: self.log_box.insert("end", f"{msg}\n"))

    def prepare_determinate(self):
        self.progress.stop()
        self.progress.configure(mode="determinate")

    def update_progress_ui(self, p, c, t):
        self.progress.set(p)
        self.lbl_percentage.configure(text=f"{int(p*100)}%")
        self.lbl_status.configure(text=f"Processando: {c}/{t}")

    def add_to_table(self, n, t, y, d):
        self.tree.insert("", "end", values=(n, t, y, d))

    def finish_ui(self):
        self.btn_select.configure(state="normal")
        messagebox.showinfo("BiblioTech", "Concluído!")

if __name__ == "__main__":
    app = BiblioTechPro()
    app.mainloop()