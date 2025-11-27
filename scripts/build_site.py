import os
import html
from datetime import datetime
from pathlib import Path

from docx import Document

# Pastas principais
WORD_DIR = Path("word-artigos")
HTML_DIR = Path("artigos")
INDEX_FILE = Path("index.html")


def ensure_directories():
    """Garante que as pastas necessárias existam."""
    HTML_DIR.mkdir(exist_ok=True)
    WORD_DIR.mkdir(exist_ok=True)


def extract_article_from_docx(docx_path: Path):
    """
    Lê um DOCX e devolve:
    - título  -> primeiro parágrafo não vazio
    - parágrafos do corpo
    - resumo (para o card no índice)
    """
    doc = Document(str(docx_path))
    paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]

    if not paragraphs:
        return None

    title = paragraphs[0]
    body_paragraphs = paragraphs[1:] or paragraphs[:1]

    # resumo simples: primeiras 40–60 palavras do corpo
    joined = " ".join(body_paragraphs)
    words = joined.split()
    summary = " ".join(words[:50])
    if len(words) > 50:
        summary += "..."

    return title, body_paragraphs, summary


def build_article_html(title: str, body_paragraphs):
    """Monta o HTML de um artigo individual."""
    body_html = "\n".join(
        f"<p>{html.escape(p)}</p>" for p in body_paragraphs
    )

    article_html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <title>{html.escape(title)} – Artigos de Moacir Lázaro Melo</title>
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <style>
    body {{
      font-family: system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      margin: 0;
      padding: 0;
      background: #f4f6fb;
      color: #1b2430;
      line-height: 1.6;
    }}
    header {{
      background: linear-gradient(135deg, #003b8e, #0058c0);
      color: #fff;
      padding: 24px 16px;
    }}
    header a {{
      color: #ffe26a;
      text-decoration: none;
      font-size: 0.9rem;
    }}
    header a:hover {{
      text-decoration: underline;
    }}
    .container {{
      max-width: 768px;
      margin: 0 auto;
      padding: 24px 16px 40px;
      background: #ffffff;
    }}
    h1 {{
      font-size: 1.8rem;
      margin-top: 0;
      color: #12223b;
    }}
    p {{
      font-size: 1rem;
      margin: 0 0 1rem;
    }}
    footer {{
      text-align: center;
      font-size: 0.8rem;
      color: #6b7280;
      padding: 16px;
    }}
  </style>
</head>
<body>
  <header>
    <div style="max-width: 768px; margin: 0 auto; display: flex; justify-content: space-between; align-items: center;">
      <div>
        <div style="font-size: 0.8rem; opacity: .9;">ROTARY • ESCRITOR • EMPRESÁRIO</div>
        <div style="font-size: 1.2rem; font-weight: 600;">Artigos de Moacir Lázaro Melo</div>
      </div>
      <a href="index.html">Voltar à biblioteca de artigos</a>
    </div>
  </header>

  <main class="container">
    <h1>{html.escape(title)}</h1>
    {body_html}
  </main>

  <footer>
    &copy; {datetime.now().year} Moacir Lázaro Melo — Biblioteca de artigos.
  </footer>
</body>
</html>
"""
    return article_html


def generate_articles():
    """
    Converte todos os DOCX em WORD_DIR para HTML em HTML_DIR
    e devolve uma lista de metadados para montar o índice.
    """
    articles_meta = []

    for docx_name in sorted(os.listdir(WORD_DIR)):
        if not docx_name.lower().endswith(".docx"):
            continue

        docx_path = WORD_DIR / docx_name
        # mantém o nome base do arquivo, apenas troca a extensão
        html_name = f"{docx_path.stem}.html"
        html_path = HTML_DIR / html_name

        extracted = extract_article_from_docx(docx_path)
        if not extracted:
            continue

        title, body_paragraphs, summary = extracted

        # gera HTML do artigo
        article_html = build_article_html(title, body_paragraphs)
        html_path.write_text(article_html, encoding="utf-8")

        # data a partir do arquivo DOCX (última modificação)
        dt = datetime.fromtimestamp(docx_path.stat().st_mtime)
        date_str = dt.strftime("%d/%m/%Y")

        articles_meta.append(
            {
                "title": title,
                "summary": summary,
                "html_file": html_name,
                "date": dt,
                "date_str": date_str,
            }
        )

    # artigos mais recentes primeiro
    articles_meta.sort(key=lambda a: a["date"], reverse=True)
    return articles_meta


def build_index_html(articles_meta):
    """Gera o conteúdo completo do index.html com base nos metadados."""

    if articles_meta:
        cards = []
        for art in articles_meta:
            cards.append(
                f"""
        <article class="card">
          <div class="card-top">
            <span class="pill pill-type">Artigo</span>
            <span class="pill pill-date">{html.escape(art['date_str'])}</span>
          </div>
          <h2 class="card-title">
            <a href="artigos/{html.escape(art['html_file'])}">
              {html.escape(art['title'])}
            </a>
          </h2>
          <p class="card-summary">{html.escape(art['summary'])}</p>
        </article>
        """
            )
        articles_html = "\n".join(cards)
    else:
        articles_html = """
        <p style="color:#6b7280;">Nenhum artigo foi encontrado ainda. Envie arquivos DOCX para a pasta <strong>word-artigos</strong> e o site será atualizado automaticamente.</p>
        """

    index_html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <title>Artigos de Moacir Lázaro Melo</title>
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <style>
    * {{
      box-sizing: border-box;
    }}
    body {{
      margin: 0;
      font-family: system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      background: radial-gradient(circle at top left, #17427d 0, #020617 40%, #020617 100%);
      color: #f9fafb;
    }}
    .page {{
      min-height: 100vh;
      display: flex;
      flex-direction: column;
    }}
    header {{
      padding: 24px 16px 8px;
    }}
    .header-inner {{
      max-width: 1080px;
      margin: 0 auto;
      display: flex;
      flex-wrap: wrap;
      gap: 24px;
      align-items: center;
      justify-content: space-between;
    }}
    .tagline {{
      font-size: 0.75rem;
      letter-spacing: .15em;
      text-transform: uppercase;
      color: #e5e7eb;
      opacity: .85;
    }}
    .title {{
      font-size: 2rem;
      font-weight: 700;
    }}
    .title span.highlight {{
      color: #ffe26a;
    }}
    .subtitle {{
      margin-top: 8px;
      max-width: 520px;
      color: #e5e7eb;
      font-size: 0.95rem;
    }}
    .bio-box {{
      display: flex;
      align-items: center;
      gap: 16px;
      padding: 12px 16px;
      border-radius: 999px;
      background: rgba(15,23,42,.85);
      margin-top: 16px;
    }}
    .bio-photo {{
      width: 64px;
      height: 64px;
      border-radius: 999px;
      overflow: hidden;
      border: 2px solid #ffe26a;
      flex-shrink: 0;
    }}
    .bio-photo img {{
      width: 100%;
      height: 100%;
      object-fit: cover;
    }}
    .bio-text-main {{
      font-size: 0.9rem;
      font-weight: 600;
      color: #e5e7eb;
    }}
    .bio-text-sub {{
      font-size: 0.8rem;
      color: #9ca3af;
    }}
    .hero-badge {{
      font-size: 0.78rem;
      color: #c7d2fe;
      margin-top: 4px;
    }}
    main {{
      flex: 1;
      padding: 0 16px 32px;
    }}
    .content {{
      max-width: 1080px;
      margin: 0 auto;
      background: #0b1220;
      border-radius: 24px 24px 0 0;
      margin-top: 16px;
      padding: 24px 16px 32px;
    }}
    .filters-row {{
      display: flex;
      flex-wrap: wrap;
      gap: 12px;
      margin-bottom: 20px;
    }}
    .search-input {{
      flex: 1;
      min-width: 200px;
      padding: 10px 12px;
      border-radius: 999px;
      border: 1px solid #1e293b;
      background: #020617;
      color: #f9fafb;
      font-size: 0.9rem;
    }}
    .search-input::placeholder {{
      color: #64748b;
    }}
    .articles-grid {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(260px, 1fr));
      gap: 16px;
    }}
    .card {{
      background: #020617;
      border-radius: 18px;
      padding: 16px 16px 18px;
      border: 1px solid rgba(148,163,184,.3);
      display: flex;
      flex-direction: column;
      gap: 6px;
    }}
    .card-top {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      font-size: 0.75rem;
    }}
    .pill {{
      border-radius: 999px;
      padding: 4px 10px;
    }}
    .pill-type {{
      background: rgba(56,189,248,.15);
      color: #7dd3fc;
    }}
    .pill-date {{
      background: rgba(148,163,184,.18);
      color: #e5e7eb;
    }}
    .card-title {{
      font-size: 1rem;
      margin: 0;
    }}
    .card-title a {{
      color: #e5e7eb;
      text-decoration: none;
    }}
    .card-title a:hover {{
      color: #ffe26a;
    }}
    .card-summary {{
      font-size: 0.85rem;
      color: #9ca3af;
      margin: 0;
    }}
    footer {{
      text-align: center;
      font-size: 0.75rem;
      color: #9ca3af;
      padding: 16px;
    }}
    @media (min-width: 768px) {{
      .title {{
        font-size: 2.4rem;
      }}
      header {{
        padding-top: 32px;
      }}
      .content {{
        padding: 28px 24px 36px;
      }}
    }}
  </style>
</head>
<body>
<div class="page">
  <header>
    <div class="header-inner">
      <div>
        <div class="tagline">ROTARY • ESCRITOR • EMPRESÁRIO</div>
        <div class="title">
          Artigos de <span class="highlight">Moacir Lázaro Melo</span>
        </div>
        <p class="subtitle">
          Textos publicados em jornais e revistas, reunidos aqui para leitura completa.
          Reflexões sobre o país, gerações, convivência humana, serviço ao próximo e o sentido de viver bem.
        </p>
        <div class="bio-box">
          <div class="bio-photo">
            <img src="assets/moacir-melo.jpg" alt="Foto de Moacir Lázaro Melo">
          </div>
          <div>
            <div class="bio-text-main">Moacir Lázaro Melo</div>
            <div class="bio-text-sub">
              Escritor, empresário e rotariano com décadas dedicadas a servir comunidades e inspirar pessoas.
            </div>
            <div class="hero-badge">
              “Melhorar vidas via Rotary Internacional.”
            </div>
          </div>
        </div>
      </div>
    </div>
  </header>

  <main>
    <section class="content">
      <div class="filters-row">
        <input class="search-input" type="search" placeholder="Pesquisar por título ou palavra-chave (busca visual – use o CTRL+F do navegador para pesquisa completa)...">
      </div>

      <section class="articles-grid">
        {articles_html}
      </section>
    </section>
  </main>

  <footer>
    &copy; {datetime.now().year} Moacir Lázaro Melo — Biblioteca de artigos. Desenvolvido para leitura tranquila e organizada.
  </footer>
</div>
</body>
</html>
"""
    return index_html


def main():
    ensure_directories()
    articles_meta = generate_articles()
    index_html = build_index_html(articles_meta)
    INDEX_FILE.write_text(index_html, encoding="utf-8")


if __name__ == "__main__":
    main()
