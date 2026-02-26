from pathlib import Path
import shutil
import logging


def gerar_nome_bkp_incremental(pasta_bkp: Path, nome_base: str) -> Path:
    contador = 1

    while True:
        novo_nome = f"{nome_base}_{contador}.xlsx"
        destino = pasta_bkp / novo_nome

        if not destino.exists():
            return destino

        contador += 1


def mover_para_bkp(
    caminho_arquivo: Path,
    pasta_bkp: Path,
    logger: logging.Logger
) -> None:

    if not caminho_arquivo.exists():
        logger.info("Nenhum arquivo anterior encontrado para backup.")
        return

    pasta_bkp.mkdir(parents=True, exist_ok=True)

    nome_base = caminho_arquivo.stem
    novo_destino = gerar_nome_bkp_incremental(pasta_bkp, nome_base)

    logger.info(f"Iniciando backup de {caminho_arquivo.name}")
    shutil.move(str(caminho_arquivo), str(novo_destino))
    logger.info(f"Backup criado: {novo_destino}")