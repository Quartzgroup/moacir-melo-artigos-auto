import os
import re
import mammoth
from datetime import datetime

# Caminhos das pastas
WORD_DIR = "word-artigos"
HTML_DIR = "artigos"

# Template da página de artigo
ARTICLE_TEMPLATE = """
<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<title>{titulo}</title>
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<style>
body {{
    font-family: Arial, sans-serif;
    max-width: 800px;
    margin: 40px auto;
    padding: 20px;
    line-height: 1.8;
}}
h1 {{
    text-align: center;
    color: #003366;
}}
</style>
</head>
<body>

<h1>{titulo}</h1>
<p><strong>Veículo:</strong> {veiculo}</p>
<p><strong>Data:</strong> {data}</p>
<hr>

{conteudo}

</body>
</html>
"""

# Função para criar slugs de nomes
def slugify(text):
    text = text.lower()
    text = re.sub(r"[áàãâä]", "a", text)
    text = re.sub(r"[éèêë]", "e", text)
    text = re.sub(r"[íìîï]", "i", text)
    text = re.sub(r"[óòõôö]", "o", text)
    text = re.sub(r"[úùûü]", "u", text)
    text = re.sub(r"ç", "c", text)
    text = re.sub(r"[^a-z0-9]+", "-", text)
    return text.strip("-")

# Função principal para conversão
def converter_artigos():
    print("Convertendo DOCX para HTML...")

    artigos_info = []

    for filename in os.listdir(WORD_DIR):
        if filename.endswith(".docx"):
            caminho = os.path.join(WORD_DIR, filename)
            titulo = filename.replace(".docx", "").replace("-", " ").title()

            slug = slugify(titulo)
            output_html = os.path.join(HTML_DIR, f"{slug}.html")

            with open(caminho, "rb") as docx_file:
                result = mammoth.convert_to_html(docx_file)
                conteudo_html = result.value

            html_final = ARTICLE_TEMPLATE.format(
                titulo=titulo,
                conteudo=conteudo_html,
                veiculo="A definir",
                data=datetime.today().strftime("%d/%m/%Y")
            )

            with open(output_html, "w", encoding="utf-8") as html_file:
                html_file.write(html_final)

            artigos_info.append((titulo, slug))

    return artigos_info

# Gera index.html automaticamente
def gerar_index(artigos):
    print("Gerando index.html...")

    lista = ""
    for titulo, slug in artigos:
        lista += f"<li><a href='artigos/{slug}.html'>{titulo}</a></li>\n"

    INDEX_TEMPLATE = f"""
<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<title>Artigos de Moacir Lázaro Melo</title>
</head>
<body>
<h1>Artigos de Moacir Lázaro Melo</h1>
<ul>
{lista}
</ul>
</body>
</html>
"""

    with open("index.html", "w", encoding="utf-8") as f:
        f.write(INDEX_TEMPLATE)

def main():
    artigos = converter_artigos()
    gerar_index(artigos)
    print("Processo concluído!")

if __name__ == "__main__":
    main()

