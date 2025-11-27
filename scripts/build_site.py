import re
import unicodedata
from pathlib import Path

from docx import Document

ROOT = Path(".")
WORD_DIR = ROOT / "word-artigos"
HTML_DIR = ROOT / "artigos"

HTML_DIR.mkdir(exist_ok=True)

def slugify(text: str) -> str:
    text = text.strip().lower()
    text = unicodedata.normalize("NFKD", text)
    text = text.encode("ascii", "ignore").decode("ascii")
    text = re.sub(r"[^a-z0-9]+", "-", text)
    text = re.sub(r"-+", "-", text)
    return text.strip("-") or "artigo"

def docx_to_paragraphs(path: Path):
    doc = Document(str(path))
    paras = []
    for para in doc.paragraphs:
        text = para.text.strip()
        if text:
            paras.append(text)
    return paras

def escape_html(text: str) -> str:
    return (
        text.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
    )

ARTICLE_TEMPLATE = """<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <title>{titulo}</title>
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <style>
    body {{
      font-family: Arial, sans-serif;
      max-width: 900px;
      margin: 40px auto;
      padding: 0 20px 40px;
      line-height: 1.8;
      background: #f5f7fb;
      color: #0f172a;
    }}
    h1 {{
      text-align: center;
      color: #0c3c8c;
      margin-bottom: 10px;
    }}
    .meta {{
      text-align: center;
      font-size: 14px;
      color: #64748b;
      margin-bottom: 20px;
    }}
    p {{
      margin: 0 0 14px;
      text-align: justify;
    }}
    a.voltar {{
      display: inline-block;
      margin-bottom: 15px;
      text-decoration: none;
      color: #0c3c8c;
      font-weight: bold;
    }}
  </style>
</head>
<body>

<a href="../index.html" class="voltar">← Voltar para a página inicial</a>

<h1>{titulo}</h1>
<div class="meta">
  Veículo: {veiculo} • Data: {data}
</div>

{conteudo}

</body>
</html>
"""

def gerar_artigos():
    artigos = []

    for docx_path in WORD_DIR.glob("*.docx"):
        paras = docx_to_paragraphs(docx_path)
        if not paras:
            continue

        titulo = paras[0].strip()
        slug = slugify(docx_path.stem)

        corpo_html = "\n".join(f"<p>{escape_html(p)}</p>" for p in paras)

        html = ARTICLE_TEMPLATE.format(
            titulo=escape_html(titulo),
            conteudo=corpo_html,
            veiculo="Jornais e Revistas",
            data="[inserir]"
        )

        out_path = HTML_DIR / f"{slug}.html"
        out_path.write_text(html, encoding="utf-8")

        artigos.append({"titulo": titulo, "slug": slug})

    return artigos

def gerar_index(artigos):
    itens = "\n".join(
        f"<li><a href='artigos/{a['slug']}.html'>{escape_html(a['titulo'])}</a></li>"
        for a in artigos
    )

    index_html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <title>Artigos de Moacir Lázaro Melo</title>
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <style>
    body {{
      font-family: Arial, sans-serif;
      max-width: 900px;
      margin: 40px auto;
      padding: 0 20px 40px;
      background: #f5f7fb;
      color: #0f172a;
    }}
    h1 {{
      text-align: center;
      color: #0c3c8c;
      margin-bottom: 20px;
    }}
    ul {{
      list-style: none;
      padding: 0;
    }}
    li {{
      margin-bottom: 8px;
    }}
    a {{
      text-decoration: none;
      color: #0b1b3b;
      font-weight: bold;
    }}
    a:hover {{
      text-decoration: underline;
    }}
  </style>
</head>
<body>
  <h1>Artigos de Moacir Lázaro Melo</h1>
  <ul>
{itens}
  </ul>
</body>
</html>
"""
    Path("index.html").write_text(index_html, encoding="utf-8")

def main():
    artigos = gerar_artigos()
    if artigos:
        gerar_index(artigos)
    else:
        print("Nenhum DOCX encontrado em word-artigos.")

if __name__ == "__main__":
    main()

