# src/ItauSiteGov/config.py
import os
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

# =========================
# 🔐 CREDENCIAIS (ENV)
# =========================
GOVBR_CPF = os.getenv("GOVBR_CPF")
GOVBR_PASSWORD = os.getenv("GOVBR_PASSWORD")

if not GOVBR_CPF or not GOVBR_PASSWORD:
    raise RuntimeError("Credenciais GOV.BR não encontradas no .env")


# =========================
# 🌐 URLS
# =========================
URL_LOGIN = (
    "https://consumidor.gov.br/pages/administrativo/login"
    "?error=login.sem_permissao"
)

URL_RELATORIO_GERENCIAL = "/pages/relatorio/gerencial/abrir"


# =========================
# 📥 DOWNLOAD
# =========================
DOWNLOAD_DIR = Path.home() / "Downloads"


# =========================
# 📂 DESTINOS NA REDE
# =========================
BASE_NETWORK_PATH = Path(
    r"\\BRSBESRV960\publico\REPORTS\001 CARGAS\001 EXCEL\ITAU\GOV"
)

DEST_ITAU_CONSIGNADO = BASE_NETWORK_PATH / "IC"
DEST_ITAU_UNIBANCO_CONSIGNADO = BASE_NETWORK_PATH / "PF"


# =========================
# 🏦 FORNECEDORES (IDs DO SITE)
# =========================
FORNECEDORES = {
    "ITAU_CONSIGNADO": {
        "label": "Itaú Consignado",
        "value": "snniq8lkf3pTFO111Dbgn1UEReubMCud",
        "dest_dir": DEST_ITAU_CONSIGNADO,
    },
    "ITAU_UNIBANCO_CONSIGNADO": {
        "label": "Itaú Unibanco Consignado",
        "value": "snniq8lkf3pTFO111Dbgn0GNhS4Boao6",
        "dest_dir": DEST_ITAU_UNIBANCO_CONSIGNADO,
    },
}


# =========================
# 📄 EXPORTAÇÃO
# =========================
EXPORT_ALL_COLUMNS_CHECKBOX_ID = "colunasExportadas1"
EXPORT_BUTTON_ID = "btnExportar"


# =========================
# 📅 REGRAS DE NEGÓCIO
# =========================
MESES_POR_EXTRACAO = 2
