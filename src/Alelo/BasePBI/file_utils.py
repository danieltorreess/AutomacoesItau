import shutil
from pathlib import Path


def ajustar_nome_arquivo(caminho_arquivo):
    """
    Remove '_atento' do nome do arquivo
    """

    caminho = Path(caminho_arquivo)

    if not caminho.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho}")

    novo_nome = caminho.name.replace("_atento", "")
    novo_caminho = caminho.parent / novo_nome

    caminho.rename(novo_caminho)

    return novo_caminho


def copiar_para_destinos(caminho_arquivo, destinos):

    caminho_arquivo = Path(caminho_arquivo)

    for destino in destinos:
        destino_path = Path(destino)
        destino_path.mkdir(parents=True, exist_ok=True)

        shutil.copy2(caminho_arquivo, destino_path)

        print(f"📂 Copiado para: {destino_path}")
