from pathlib import Path
import time
import traceback

from src.Alelo.Projuris.processor import (
    encontrar_arquivo,
    salvar_base
)

from src.Alelo.Projuris.file_utils import mover_para_bkp
from src.Alelo.Projuris.logger import configurar_logger
from src.Alelo.Projuris.config import (
    DOWNLOAD_DIR,
    ARQUIVO_ORIGEM,
    DESTINO_REDE,
    PASTA_BKP,
    NOME_FINAL,
    PASTA_LOG
)


def main():

    logger = configurar_logger("RPA_Projuris", PASTA_LOG)

    logger.info("🚀 Iniciando RPA Projuris")

    try:

        # 🔎 Localizar arquivo
        arquivo_origem = encontrar_arquivo(
            DOWNLOAD_DIR,
            ARQUIVO_ORIGEM
        )

        if not arquivo_origem:
            logger.warning("Arquivo não encontrado na pasta Downloads.")
            return

        logger.info(f"Arquivo encontrado: {arquivo_origem}")

        caminho_destino = DESTINO_REDE / NOME_FINAL

        # 📦 Backup
        mover_para_bkp(
            caminho_destino,
            PASTA_BKP,
            logger
        )

        # 💾 Salvar nova base
        salvar_base(
            arquivo_origem,
            caminho_destino,
            logger
        )

        # 🗑 Remover original
        logger.info("Removendo arquivo da pasta Downloads...")

        for tentativa in range(3):
            try:
                time.sleep(1)
                arquivo_origem.unlink()
                logger.info("Arquivo removido com sucesso.")
                break
            except PermissionError:
                logger.warning(
                    f"Arquivo em uso. Tentativa {tentativa + 1}/3"
                )
                time.sleep(2)
        else:
            logger.error("Não foi possível remover o arquivo.")

        logger.info("✅ Processo finalizado com sucesso.")

    except Exception:
        logger.error("Erro inesperado no processo.")
        logger.error(traceback.format_exc())


if __name__ == "__main__":
    main()