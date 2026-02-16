import re
import html


def extrair_link_download(html_body):
    """
    Extrai o link completo do Azure Blob e
    corrige entidades HTML (&amp; → &)
    """

    if not html_body:
        return None

    html_limpo = html_body.replace("\r", "").replace("\n", "")

    padrao = r"(https://stlandzoneprd001\.blob\.core\.windows\.net[^\s\"'<]+)"
    match = re.search(padrao, html_limpo)

    if match:
        link = match.group(1)
        link_corrigido = html.unescape(link)
        return link_corrigido

    return None
