from scholarly import scholarly

def pesquisar_no_scholar(termo_pesquisa):
    print(f"Pesquisando por: '{termo_pesquisa}'...")
    
    # Faz a busca
    search_query = scholarly.search_pubs(termo_pesquisa)
    
    # Pega o primeiro resultado da busca
    try:
        primeiro_resultado = next(search_query)
        
        # Exibe os dados do resultado
        print("\n--- Resultado Encontrado ---")
        print(f"Título: {primeiro_resultado['bib']['title']}")
        print(f"Autores: {primeiro_resultado['bib']['author']}")
        print(f"Ano: {primeiro_resultado['bib']['pub_year']}")
        print(f"Citações: {primeiro_resultado['num_citations']}")
        
        # Gera a citação em formato BibTeX
        print("\n--- Código BibTeX ---")
        print(scholarly.bibtex(primeiro_resultado))
        
    except StopIteration:
        print("Nenhum resultado encontrado.")

# Executar a pesquisa
pesquisar_no_scholar('inteligência artificial na medicina')