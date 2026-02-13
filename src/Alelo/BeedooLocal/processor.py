import pandas as pd
from pathlib import Path


DOWNLOAD_DIR = Path.home() / "Downloads"


def encontrar_arquivo(prefixo: str):

    arquivos = list(DOWNLOAD_DIR.glob(f"{prefixo}_*.xlsx"))

    if not arquivos:
        return None

    # pega o mais recente
    arquivos.sort(key=lambda x: x.stat().st_mtime, reverse=True)

    return arquivos[0]


def processar_excel_para_csv(caminho_xlsx: Path, nome_final: str, destino_final: Path):

    print(f"🔄 Processando {caminho_xlsx.name}")

    df = pd.read_excel(caminho_xlsx, dtype=str)

    # Salva temporariamente com sheet correta (apenas controle lógico)
    caminho_csv = destino_final / f"{nome_final}.csv"

    df.to_csv(caminho_csv, index=False, encoding="utf-8-sig")

    print(f"✅ CSV gerado: {caminho_csv}")

    return caminho_csv
