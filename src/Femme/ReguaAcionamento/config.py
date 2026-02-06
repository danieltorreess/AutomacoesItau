from pathlib import Path

URL_LOGIN = "https://itsmdigital.likewatercs.net/login"

DOWNLOAD_DIR = Path.home() / "Downloads"

DEST_DIR = Path(
    r"\\brsbesrv960\Publico\REPORTS\FEMME\DESEMPENHO\D-1\REGUA_ACIONAMENTO"
)

BKP_DIR = DEST_DIR / "BKP"

EXPORT_FILENAME = "relatorio_acionamentos.csv"
