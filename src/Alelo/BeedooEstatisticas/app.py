import os
from pathlib import Path
from dotenv import load_dotenv

from src.Alelo.BeedooEstatisticas.browser_edge import get_browser
from src.Alelo.BeedooEstatisticas.bot import BeedooEstatisticasBot
from src.Alelo.BeedooEstatisticas import config


def main():

    print("🚀 Iniciando RPA - Beedoo Estatísticas")

    # ==========================
    # 🔐 Variáveis de ambiente
    # ==========================
    root_dir = Path(__file__).resolve().parents[3]
    load_dotenv(root_dir / ".env")

    email = os.getenv("ALELO_BEEDOO_EMAIL")
    password = os.getenv("ALELO_BEEDOO_PASSWORD")

    if not all([email, password]):
        raise RuntimeError("❌ Variáveis ALELO_BEEDOO não configuradas no .env")

    # ==========================
    # 🌐 Browser
    # ==========================
    driver = get_browser(config.DOWNLOAD_DIR)

    bot = BeedooEstatisticasBot(driver, email, password, config)

    # ==========================
    # 🔐 Login
    # ==========================
    print("🔐 Realizando login...")
    bot.login()

    # ==========================
    # 📊 Navegar até Estatísticas
    # ==========================
    print("📊 Abrindo Estatísticas...")
    bot.abrir_estatisticas()

    # ==========================
    # 📈 Abrindo Feed Analítico
    # ==========================
    print("📈 Abrindo Relatório Analítico...")
    bot.abrir_feed_analitico()

    print("✅ Fluxo executado até Relatório Analítico com sucesso!")


if __name__ == "__main__":
    main()
