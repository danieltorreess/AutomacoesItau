# type: ignore
import os
from pathlib import Path
from dotenv import load_dotenv

from src.WedukaIncidentes import config
# from src.WedukaIncidentes.browser import get_browser
from src.WedukaIncidentes.browser_edge import get_browser
from src.WedukaIncidentes.weduka_incidentes_bot import WedukaIncidentesBot


def main():
    root_dir = Path(__file__).resolve().parents[2]
    load_dotenv(root_dir / ".env")

    username = os.getenv("WEDUKA_USERNAME")
    password = os.getenv("WEDUKA_PASSWORD")

    # driver = get_browser(config.DOWNLOAD_DIR)
    driver = get_browser(config.DOWNLOAD_DIR)
    bot = WedukaIncidentesBot(driver, username, password, config)

    bot.login()
    bot.acessar_relatorio()
    bot.extrair_por_dia()

    # driver.quit()


if __name__ == "__main__":
    main()
