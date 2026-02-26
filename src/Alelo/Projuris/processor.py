from pathlib import Path
import shutil
import logging


def encontrar_arquivo(download_dir: Path, nome_arquivo: str) -> Path | None:
    caminho = download_dir / nome_arquivo

    if caminho.exists():
        return caminho

    return None


def salvar_base(
    caminho_origem: Path,
    caminho_destino: Path,
    logger: logging.Logger
) -> None:

    logger.info(f"Copiando arquivo para: {caminho_destino}")

    caminho_destino.parent.mkdir(parents=True, exist_ok=True)

    shutil.copy2(caminho_origem, caminho_destino)

    logger.info("Arquivo salvo com sucesso.")