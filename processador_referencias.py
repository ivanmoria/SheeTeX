# processador_referencias.py

import re
import pandas as pd
from PyQt5.QtWidgets import (QMessageBox)

def extrair_campos_apa(texto):
    campos = {
        'author': '', 'year': '', 'title': '',
        'journal': '', 'booktitle': '', 'number': '',
        'pages': '', 'doi': ''
    }
    texto = ' '.join(texto.split())  # tira quebras e espaços extras

    match_ano = re.search(r'\((\d{4})\)', texto)
    if not match_ano:
        return campos
    campos['year'] = match_ano.group(1)
    inicio_ano = match_ano.start()
    fim_ano = match_ano.end()

    # author antes do ano, removendo ponto final final
    campos['author'] = texto[:inicio_ano].strip().rstrip('.')

    depois = texto[fim_ano:].strip()

    # Extrai título até o primeiro ponto final
    match_titulo = re.match(r'\.?\s*([^\.]+)\.', depois)
    if match_titulo:
        campos['title'] = match_titulo.group(1).strip()
        resto = depois[match_titulo.end():].strip()
    else:
        # se não encontrar título com ponto final, pega tudo como título
        campos['title'] = depois.strip()
        resto = ""

    # Agora analisa resto para booktitle, journal, pages, etc

    # Se tem "In " no começo do resto, extrai booktitle e páginas
    if resto.startswith('In '):
        resto_sem_in = resto[3:].strip()
        # Extrai booktitle até ponto final ou (pp.
        match_booktitle = re.match(r'(.+?)(?:\.|\(pp\.)', resto_sem_in)
        if match_booktitle:
            campos['booktitle'] = match_booktitle.group(1).strip()

        match_paginas = re.search(r'\(pp\.\s*(\d+-\d+)\)', resto_sem_in)
        if match_paginas:
            campos['pages'] = match_paginas.group(1)

    else:
        # tenta extrair journal se não tiver "In "
        match_journal = re.match(r'([^.]+)\. (.+)', resto)
        if match_journal:
            campos['journal'] = match_journal.group(1).strip()
            # poderia continuar processando o que sobrou se quiser

    # número volume(issue)
    match_number = re.search(r'(\d+\s*\(\d+\))', texto)
    if match_number:
        campos['number'] = match_number.group(1)

    # páginas caso não tenham sido capturadas antes
    if not campos['pages']:
        match_pages = re.search(r'(\d{1,4}-\d{1,4})', texto)
        if match_pages:
            campos['pages'] = match_pages.group(1)

    # doi
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



def carregar_excel(self):
        # Tenta usar self.sheet_csv_url se existir e não for vazio
        url_csv = None
        if hasattr(self, 'sheet_csv_url') and self.sheet_csv_url:
            url_csv = self.sheet_csv_url
        else:
            # Caso não exista ou seja vazio, usa a url do QLineEdit
            url = self.url_edit.text().strip()
            if not url:
                url = self.url_padrao  # usa padrão se vazio
            url_csv = self.transformar_link_para_csv(url)

        try:
            df = pd.read_csv(url_csv)
            if "Ref" not in df.columns:
                QMessageBox.critical(self, "Erro", "A coluna 'Ref' não foi encontrada na planilha.")
                return
            referencias = "\n\n".join(str(ref).strip() for ref in df["Ref"] if pd.notna(ref))
            self.input_edit.setPlainText(referencias)
            self.converter()
            self.salvar_bibtex()
            self.bibtex_viewer.reload_data()
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao ler a planilha do Google Sheets:\n{e}")