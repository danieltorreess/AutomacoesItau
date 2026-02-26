from pathlib import Path
import time
import traceback

from src.Alelo.MultivanEmail.processor import (
    encontrar_arquivo,
    salvar_na_rede
)

from src.Alelo.MultivanEmail.file_utils import mover_para_bkp
from src.Alelo.MultivanEmail.logger import configurar_logger
from src.Alelo.MultivanEmail.config import (
    DOWNLOAD_DIR,
    ARQUIVO_ORIGEM,
    DESTINO_REDE,
    PASTA_BKP,
    NOME_FINAL,
    PASTA_LOG
)


def main():

    logger = configurar_logger("RPA_MultivanEmail", PASTA_LOG)

    logger.info("🚀 Iniciando RPA MultivanEmail")

    try:

        # 🔎 Localiza arquivo
        arquivo_origem = encontrar_arquivo(
            DOWNLOAD_DIR,
            ARQUIVO_ORIGEM
        )

        if not arquivo_origem:
            logger.warning("Arquivo não encontrado na pasta Downloads.")
            return

        logger.info(f"Arquivo encontrado: {arquivo_origem}")

        caminho_destino = DESTINO_REDE / NOME_FINAL

        # 📦 Backup anterior
        mover_para_bkp(
            caminho_destino,
            PASTA_BKP,
            logger
        )

        # 💾 Copia para rede
        salvar_na_rede(
            arquivo_origem,
            caminho_destino,
            logger
        )

        # 🗑 Remove original
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