from pathlib import Path

# URLs
URL_INTEGRATION = "https://base.weduka.com.br/Account/Integration"

# Diretórios
DOWNLOAD_DIR = Path.home() / "Downloads"
# DEST_DIR = Path(r"C:\ItauWedukaTreinamentos\Procedimentos")
DEST_DIR = Path(r"\\BRSBESRV960\Publico\REPORTS\ITAU\WEDUKA_TREINAMENTO\Procedimentos")

# Repositórios a extrair
REPOSITORIOS = [
    "Banco Pessoa Física",
    "Empresas",
    "Cartões",
    "Geral Negócios",
    "Financiamento de Veículos",
    "Consignado PF",
    "Rede",
    "NPC",
    "Backoffice Cartões",
    "Cartões PJ",
]

# Prefixo padrão (SSIS vai consumir isso)
FILE_PREFIX = "download_procedimentos_"
