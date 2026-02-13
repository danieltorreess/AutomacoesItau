import requests
from pathlib import Path
from urllib.parse import urlparse


def baixar_arquivo(url, pasta_download):

    response = requests.get(
        url,
        stream=True,
        allow_redirects=True,
        timeout=60
    )

    if response.status_code != 200:
        raise Exception(f"Erro ao baixar arquivo: {response.status_code}")

    # 🔥 Extrai nome corretamente da URL original
    parsed_url = urlparse(url)
    nome_arquivo = Path(parsed_url.path).name

    if not nome_arquivo.endswith(".xlsx"):
        raise Exception("Nome do arquivo inválido detectado na URL.")

    caminho_final = Path(pasta_download) / nome_arquivo

    with open(caminho_final, "wb") as f:
        for chunk in response.iter_content(chunk_size=8192):
            if chunk:
                f.write(chunk)

    return caminho_final
