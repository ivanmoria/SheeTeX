import bibtexparser
import pandas as pd
import re
from collections import Counter
from habanero import Crossref
from tqdm import tqdm

# =========================
# CONFIG
# =========================
bib_file = "references.bib"
output_excel = "analysis_report.xlsx"

cr = Crossref()

# =========================
# FUNÇÕES DE LIMPEZA
# =========================

def clean_text(text):
    if not text:
        return ""
    text = re.sub(r"\s+", " ", text)
    text = text.replace("{", "").replace("}", "")
    return text.strip()

def normalize_authors(author_str):
    if not author_str:
        return ""
    authors = [a.strip() for a in author_str.split(" and ")]
    return "; ".join(authors)

def fetch_title_from_doi(doi):
    try:
        result = cr.works(ids=doi)
        return result['message']['title'][0]
    except:
        return ""

# =========================
# CARREGAR BIB
# =========================
with open(bib_file, encoding="utf-8") as bibtex_file:
    bib_db = bibtexparser.load(bibtex_file)

entries = bib_db.entries

cleaned = []

print("\n🧹 Limpando e enriquecendo dados...\n")

for e in tqdm(entries):

    title = clean_text(e.get("title", ""))
    doi = e.get("doi", "")

    # tentar recuperar título se estiver vazio
    if not title and doi:
        title = fetch_title_from_doi(doi)

    entry = {
        "ID": e.get("ID", ""),
        "TYPE": e.get("ENTRYTYPE", ""),
        "TITLE": title,
        "AUTHOR": normalize_authors(e.get("author", "")),
        "YEAR": str(e.get("year", "")),
        "JOURNAL": clean_text(e.get("journal", e.get("booktitle", ""))),
        "DOI": doi,
        "URL": e.get("url", "")
    }

    cleaned.append(entry)

df = pd.DataFrame(cleaned)

# =========================
# REMOVER DUPLICATAS
# =========================
df = df.drop_duplicates(subset=["DOI", "TITLE"])

print(f"\n📚 Total após limpeza: {len(df)}")

# =========================
# ANÁLISES
# =========================

# Artigos por ano
per_year = df["YEAR"].value_counts().sort_index()

# Autores
all_authors = []
for authors in df["AUTHOR"]:
    if authors:
        all_authors.extend(authors.split("; "))

top_authors = Counter(all_authors).most_common(20)

# Revistas
top_journals = df["JOURNAL"].value_counts().head(20)

# Palavras nos títulos
stopwords = {"the","and","of","in","for","a","an","to","on","with","from","by","at","is","are","as","this","that","be"}
words = []

for t in df["TITLE"]:
    tokens = re.findall(r"\b\w+\b", t.lower())
    tokens = [w for w in tokens if w not in stopwords and len(w) > 3]
    words.extend(tokens)

top_words = Counter(words).most_common(30)

# =========================
# EXPORTAR EXCEL
# =========================

with pd.ExcelWriter(output_excel) as writer:

    df.to_excel(writer, sheet_name="cleaned_data", index=False)

    per_year.rename_axis("YEAR").reset_index(name="count").to_excel(writer, sheet_name="per_year", index=False)

    pd.DataFrame(top_authors, columns=["Author","Count"]).to_excel(writer, sheet_name="top_authors", index=False)

    top_journals.rename_axis("Journal").reset_index(name="Count").to_excel(writer, sheet_name="top_journals", index=False)

    pd.DataFrame(top_words, columns=["Word","Count"]).to_excel(writer, sheet_name="title_words", index=False)

print("\n✅ Relatório salvo como:", output_excel)