from pathlib import Path

DOWNLOAD_DIR = Path.home() / "Downloads"

ARQUIVO_ORIGEM = "CAC VAREJO.xlsx"  # ajuste se necessário

DESTINO_REDE = Path(
    r"\\brsbesrv960\publico\REPORTS\ALELO\CAC - EC VAREJO"
)

PASTA_BKP = DESTINO_REDE / "BKP" / "VAREJO"

NOME_FINAL = "CAC_VAREJO.csv"

PASTA_LOG = DESTINO_REDE / "logs"