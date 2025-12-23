# type:ignore
import os
from pathlib import Path
from dotenv import load_dotenv

from src.WedukaTreinamento import config
from src.WedukaTreinamento.browser import get_browser
from src.WedukaTreinamento.weduka_bot import WedukaBot


def main():
    root_dir = Path(__file__).resolve().parents[2]
    env_path = root_dir / ".env"
    load_dotenv(env_path)

    username = os.getenv("WEDUKA_USERNAME")
    password = os.getenv("WEDUKA_PASSWORD")

    print(f"[WEDUKA] .env carregado de: {env_path}")
    print(f"[WEDUKA] Usuário identificado")

    driver = get_browser(config.DOWNLOAD_DIR)
    bot = WedukaBot(driver, username, password, config)

    bot.login()
    bot.acessar_relatorio()

    for repo in config.REPOSITORIOS:
        bot.extrair_repositorio(repo)

    # driver.quit()  # deixe comentado se quiser ver o navegador


if __name__ == "__main__":
    main()
