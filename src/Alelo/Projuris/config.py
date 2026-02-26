from pathlib import Path

DOWNLOAD_DIR = Path.home() / "Downloads"

ARQUIVO_ORIGEM = "CONTROLE_DIARIO_Projuris.xlsx"

DESTINO_REDE = Path(
    r"\\brsbesrv960\publico\REPORTS\ALELO\PROJURIS\base"
)

PASTA_BKP = DESTINO_REDE / "BKP"

NOME_FINAL = "CONTROLE_DIARIO_Projuris.xlsx"

PASTA_LOG = DESTINO_REDE / "logs"