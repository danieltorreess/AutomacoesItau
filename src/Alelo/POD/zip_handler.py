import zipfile
from pathlib import Path


def extrair_zip(caminho_zip: Path, pasta_destino: Path):

    with zipfile.ZipFile(caminho_zip, 'r') as zip_ref:
        zip_ref.extractall(pasta_destino)

    # Retorna pasta criada (mesmo nome do zip sem .zip)
    nome_pasta = caminho_zip.stem
    return pasta_destino / nome_pasta
