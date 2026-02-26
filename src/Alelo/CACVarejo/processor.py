# type: ignore

import pandas as pd
from pathlib import Path
import logging


def encontrar_arquivo(download_dir: Path, nome_arquivo: str) -> Path | None:
    caminho = download_dir / nome_arquivo

    if caminho.exists():
        return caminho

    return None


def processar_e_salvar_csv(
    caminho_xlsx: Path,
    caminho_destino_csv: Path,
    logger: logging.Logger
) -> None:

    logger.info(f"🔄 Processando {caminho_xlsx.name}")

    # 🔥 Leitura simples, sem alterar encoding
    df = pd.read_excel(
        caminho_xlsx,
        dtype=str
    )

    # Garante diretório
    caminho_destino_csv.parent.mkdir(parents=True, exist_ok=True)

    # 🔥 MESMO PADRÃO DO BEEDOO
    df.to_csv(
        caminho_destino_csv,
        index=False,
        sep=";",                 # padrão BR
        encoding="utf-8-sig"     # Excel friendly
    )

    logger.info(f"✅ CSV gerado: {caminho_destino_csv}")