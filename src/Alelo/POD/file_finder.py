from pathlib import Path


def encontrar_zip_mais_recente(pasta_downloads: str):
    pasta = Path(pasta_downloads)

    arquivos = list(pasta.glob("desk-messages-*.zip"))

    if not arquivos:
        return None

    # Ordena pelo mais recente (data de modificação)
    arquivos.sort(key=lambda x: x.stat().st_mtime, reverse=True)

    return arquivos[0]
