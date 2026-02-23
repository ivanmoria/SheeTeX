import asyncio
import os
from playwright.async_api import async_playwright

async def salvar_site_a4_perfeito(url, nome_arquivo):
    pasta_downloads = os.path.join(os.path.expanduser("~"), "Downloads")
    caminho_final = os.path.join(pasta_downloads, nome_arquivo)

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        # Emulamos uma tela de Desktop padrão
        context = await browser.new_context(
            viewport={'width': 1280, 'height': 800},
            device_scale_factor=1
        )
        page = await context.new_page()
        
        print(f"Acessando {url}...")
        await page.goto(url, wait_until="networkidle")

        # Roda o scroll para carregar elementos dinâmicos (Lazy Load)
        await page.evaluate("""
            async () => {
                await new Promise((resolve) => {
                    let totalHeight = 0;
                    let distance = 200;
                    let timer = setInterval(() => {
                        let scrollHeight = document.body.scrollHeight;
                        window.scrollBy(0, distance);
                        totalHeight += distance;
                        if(totalHeight >= scrollHeight){
                            clearInterval(timer);
                            resolve();
                        }
                    }, 100);
                });
            }
        """)
        
        await asyncio.sleep(2) # Pausa para renderização final

        # Ajuste de CSS para evitar quebras de imagens e textos ao meio
        await page.add_style_tag(content="""
            header, footer, section, img, tr, li { 
                break-inside: avoid !important; 
            }
        """)

        print(f"Gerando PDF...")
        await page.pdf(
            path=caminho_final,
            format="A4",
            print_background=True,
            scale=0.75,  # Reduz a escala para o layout de 1280px caber na largura do A4 (210mm)
            margin={
                "top": "1cm",
                "right": "1cm",
                "bottom": "1cm",
                "left": "1cm"
            },
            prefer_css_page_size=False 
        )
        
        await browser.close()
        print(f"PDF salvo com sucesso em: {caminho_final}")

if __name__ == "__main__":
    asyncio.run(salvar_site_a4_perfeito("https://wcmt2026.org/", "wcmt2026_corrigido.pdf"))