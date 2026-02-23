from pypdf import PdfReader, PdfWriter, Transformation

def uniformizar_pdf(input_path, output_path):
    reader = PdfReader(input_path)
    writer = PdfWriter()
    
    # 1. Pegamos APENAS a largura da primeira página (Alvo)
    target_page = reader.pages[0]
    target_width = target_page.mediabox.width
    
    for i, page in enumerate(reader.pages):
        if i == 0:
            # A primeira página já está no tamanho certo
            writer.add_page(page)
        else:
            # 2. Calculamos o fator de escala apenas para a largura (Eixo X)
            scale_x = float(target_width) / float(page.mediabox.width)
            
            # A altura recebe fator 1.0 (ou seja, não muda)
            scale_y = 1.0 
            
            # Aplicamos a transformação
            op = Transformation().scale(sx=scale_x, sy=scale_y)
            page.add_transformation(op)
            
            # 3. Ajustamos a "caixa" da página para a nova largura.
            # Não mexemos no "top" nem no "bottom" para manter a altura original.
            page.mediabox.right = target_width
            
            writer.add_page(page)

    with open(output_path, "wb") as f:
        writer.write(f)
    print(f"Sucesso! Arquivo salvo em: {output_path}")

# Caminhos do seu computador
caminho_original = "/Users/ivanmoria/Downloads/scholarship_dinner.pdf"
caminho_final = "/Users/ivanmoria/Downloads/scholarship_dinner_unifi2ca2wd1o.pdf"

uniformizar_pdf(caminho_original, caminho_final)