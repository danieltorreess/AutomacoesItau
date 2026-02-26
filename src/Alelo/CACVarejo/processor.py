#type:ignore
import pandas as pd
from pathlib import Path
import logging


def encontrar_arquivo(download_dir: Path, nome_arquivo: str) -> Path | None:
    caminho = download_dir / nome_arquivo

    if caminho.exists():
        return caminho

    return None


def processar_e_salvar_csv(
    caminho_origem: Path,
    caminho_destino: Path,
    logger: logging.Logger
) -> None:

    logger.info(f"Lendo arquivo Excel: {caminho_origem}")

    # Lê tudo como string para evitar conversão automática
    df = pd.read_excel(
        caminho_origem,
        dtype=str
    )

    logger.info(f"Linhas carregadas: {len(df)}")

    caminho_destino.parent.mkdir(parents=True, exist_ok=True)

    logger.info("Salvando CSV no padrão Excel Brasil")

    with open(caminho_destino, "w", encoding="utf-8-sig", newline="") as f:
        f.write("sep=;\n")
        df.to_csv(
            f,
            sep=";",
            index=False,
            line_terminator="\r\n"
        )

    logger.info("CSV salvo com sucesso.")