# type: ignore
import os
import time
import shutil
from pathlib import Path
from dotenv import load_dotenv
from datetime import datetime, timedelta
import pandas as pd

from src.Femme.ReguaAcionamento import config
from src.Femme.ReguaAcionamento.browser_edge import get_browser
from src.Femme.ReguaAcionamento.regua_acionamento_bot import ReguaAcionamentoBot
from src.Femme.ReguaAcionamento.utils import mover_para_bkp


def main():

    print("🚀 [START] Iniciando RPA - Régua de Acionamento FEMME")

    # ==================================================
    # 🔧 Carregar variáveis de ambiente
    # ==================================================
    root_dir = Path(__file__).resolve().parents[3]
    load_dotenv(root_dir / ".env")

    email = os.getenv("FEMME_EMAIL")
    password = os.getenv("FEMME_PASSWORD")
    tenant = os.getenv("FEMME_TENANT")

    if not all([email, password, tenant]):
        raise RuntimeError("❌ Variáveis de ambiente FEMME não configuradas corretamente")

    print("🔐 [ENV] Variáveis de ambiente carregadas")

    # ==================================================
    # 🌐 Abrir navegador
    # ==================================================
    print("🌐 [BROWSER] Abrindo Microsoft Edge")
    driver = get_browser(config.DOWNLOAD_DIR)

    bot = ReguaAcionamentoBot(driver, email, password, tenant, config)

    # ==================================================
    # 🔐 Login
    # ==================================================
    print("🔐 [LOGIN] Realizando login")
    bot.login()

    # ==================================================
    # 📊 Navegação
    # ==================================================
    print("📊 [NAV] Acessando relatório")
    bot.acessar_relatorio()

    # ==================================================
    # 📅 Extração dia a dia (D-5 até D-1)
    # ==================================================
    print("📅 [EXTRACT] Extraindo últimos 5 dias (1 dia por vez)")

    hoje = datetime.today()
    datas = [hoje - timedelta(days=i) for i in range(5, 0, -1)]

    arquivos_temporarios = []

    for data in datas:

        print(f"📅 Extraindo {data.date()}")

        bot.aplicar_filtro_e_exportar(data, data)

        origem = config.DOWNLOAD_DIR / config.EXPORT_FILENAME

        timeout = 120
        inicio = time.time()

        while not origem.exists():
            if time.time() - inicio > timeout:
                raise TimeoutError("❌ Timeout aguardando download")
            time.sleep(1)

        novo_nome = f"relatorio_{data.strftime('%Y%m%d')}.csv"
        novo_caminho = config.DOWNLOAD_DIR / novo_nome

        shutil.move(str(origem), str(novo_caminho))

        arquivos_temporarios.append(novo_caminho)

    # ==================================================
    # 📊 Consolidação dos arquivos
    # ==================================================
    print("📊 [CONSOLIDACAO] Empilhando arquivos")

    dfs = []

    for arquivo in arquivos_temporarios:
        df = pd.read_csv(arquivo, sep=';', encoding='utf-8')
        dfs.append(df)

    df_final = pd.concat(dfs, ignore_index=True)

    arquivo_consolidado = config.DOWNLOAD_DIR / config.EXPORT_FILENAME

    df_final.to_csv(
        arquivo_consolidado,
        sep=';',
        index=False,
        encoding='utf-8'
    )

    print("✅ Consolidado criado com sucesso")

    # ==================================================
    # 🧹 Remover temporários
    # ==================================================
    for arquivo in arquivos_temporarios:
        arquivo.unlink()

    # ==================================================
    # 📁 Garantir diretórios
    # ==================================================
    config.DEST_DIR.mkdir(parents=True, exist_ok=True)
    config.BKP_DIR.mkdir(parents=True, exist_ok=True)

    # ==================================================
    # ♻️ Backup do arquivo anterior
    # ==================================================
    mover_para_bkp(
        config.DEST_DIR,
        config.BKP_DIR,
        config.EXPORT_FILENAME
    )

    # ==================================================
    # 🚚 Mover consolidado para rede
    # ==================================================
    destino = config.DEST_DIR / config.EXPORT_FILENAME

    shutil.move(str(arquivo_consolidado), str(destino))

    print(f"📁 [MOVE] Arquivo salvo em: {destino}")

    print("✅ [END] RPA finalizado com sucesso")

    # driver.quit()  # opcional


if __name__ == "__main__":
    main()