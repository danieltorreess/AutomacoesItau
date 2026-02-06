# type: ignore
import os
import time
import shutil
from pathlib import Path
from dotenv import load_dotenv

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
    print("🔐 [LOGIN] Inserindo credenciais e acessando sistema")
    bot.login()
    print("🔐 [LOGIN] Login realizado com sucesso")

    # ==================================================
    # 📊 Navegação até relatório
    # ==================================================
    print("📊 [NAV] Acessando Relatório → Acionamentos")
    bot.acessar_relatorio()
    print("📊 [NAV] Relatório de Acionamentos aberto")

    # ==================================================
    # 📅 Filtro, atualização e exportação
    # ==================================================
    print("📅 [FILTRO] Setando período (D-5 até D-1), atualizando relatório e exportando CSV")
    bot.aplicar_filtro_e_exportar()

    # ==================================================
    # ⏳ Aguardar download do arquivo
    # ==================================================
    print("⬇️ [DOWNLOAD] Aguardando download do arquivo relatorio_acionamentos.csv")

    origem = config.DOWNLOAD_DIR / config.EXPORT_FILENAME
    timeout = 120  # segundos
    inicio = time.time()

    while not origem.exists():
        if time.time() - inicio > timeout:
            raise TimeoutError("❌ Timeout aguardando download do relatorio_acionamentos.csv")
        time.sleep(1)

    print(f"⬇️ [DOWNLOAD] Arquivo baixado com sucesso: {origem}")

    # ==================================================
    # 📁 Garantir diretórios de destino
    # ==================================================
    config.DEST_DIR.mkdir(parents=True, exist_ok=True)
    config.BKP_DIR.mkdir(parents=True, exist_ok=True)

    print(f"📁 [DIR] Diretório destino verificado: {config.DEST_DIR}")
    print(f"📁 [DIR] Diretório BKP verificado: {config.BKP_DIR}")

    # ==================================================
    # ♻️ Backup do arquivo anterior (se existir)
    # ==================================================
    print("♻️ [BKP] Verificando existência de arquivo anterior para backup")
    mover_para_bkp(
        config.DEST_DIR,
        config.BKP_DIR,
        config.EXPORT_FILENAME
    )

    # ==================================================
    # 🚚 Mover novo arquivo para a rede
    # ==================================================
    destino = config.DEST_DIR / config.EXPORT_FILENAME
    print(f"📁 [MOVE] Movendo arquivo para a rede: {destino}")

    shutil.move(str(origem), str(destino))

    print(f"📁 [MOVE] Arquivo salvo com sucesso em: {destino}")

    # ==================================================
    # ✅ Finalização
    # ==================================================
    print("✅ [END] RPA Régua de Acionamento finalizado com sucesso")

    # Se quiser fechar o navegador no final:
    # driver.quit()


if __name__ == "__main__":
    main()
