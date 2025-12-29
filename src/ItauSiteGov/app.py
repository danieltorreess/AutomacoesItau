# src/ItauSiteGov/app.py
from pathlib import Path

from .browser import get_browser
from .login import login_govbr
from .navigation import select_fornecedor, open_relatorios_gerenciais
from .config import DOWNLOAD_DIR


def main():
    driver = get_browser(DOWNLOAD_DIR)

    try:
        login_govbr(driver)

        # 🔹 Teste com Itaú Consignado
        select_fornecedor(driver, "ITAU_CONSIGNADO")
        open_relatorios_gerenciais(driver)

        print("✅ Navegação até Relatórios Gerenciais concluída")

    finally:
        pass
        # driver.quit()


if __name__ == "__main__":
    main()
