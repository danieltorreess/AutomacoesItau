from pathlib import Path

# URLs
URL_INTEGRATION = "https://base.weduka.com.br/Account/Integration"

# Diretórios
DOWNLOAD_DIR = Path.home() / "Downloads"
DEST_DIR = Path(r"\\BRSBESRV960\Publico\REPORTS\ITAU\WEDUKA_TREINAMENTO")

# Repositórios a extrair
REPOSITORIOS = [
    "Cartões",
    "Cartões PJ",
]

# Prefixo padrão (SSIS vai consumir isso)
FILE_PREFIX = "download_procedimentos_"
