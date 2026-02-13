from pathlib import Path
import time

from src.Alelo.PlanilhaAtribuicao.processor import (
    encontrar_planilha,
    processar_planilha
)
from src.Alelo.PlanilhaAtribuicao.file_utils import mover_para_bkp


DOWNLOAD_DIR = Path.home() / "Downloads"

DESTINO_REDE = Path(
    r"\\brsbesrv960\Publico\BASE\ALELO\BKO_ATENDIMENTO_ATRIBUICAO"
)

PASTA_BKP = DESTINO_REDE / "BKP"

NOME_FINAL = "PlanilhaAtribuicao.xlsx"


def main():

    print("🚀 Iniciando RPA PlanilhaAtribuicao")

    arquivo_origem = encontrar_planilha()

    if not arquivo_origem:
        print("⚠️ Nenhuma planilha encontrada na pasta Downloads.")
        return

    print(f"📥 Arquivo encontrado: {arquivo_origem.name}")

    caminho_destino = DESTINO_REDE / NOME_FINAL

    # ==============================
    # 📦 Backup do arquivo da rede
    # ==============================
    mover_para_bkp(caminho_destino, PASTA_BKP)

    # ==============================
    # 🔄 Processamento
    # ==============================
    processar_planilha(arquivo_origem, caminho_destino)

    # ==============================
    # 🗑 Remover arquivo original com retry
    # ==============================
    print("🗑 Removendo arquivo original dos Downloads...")

    for tentativa in range(3):
        try:
            time.sleep(1)
            arquivo_origem.unlink()
            print("🗑 Arquivo removido com sucesso.")
            break
        except PermissionError:
            print(f"⚠️ Arquivo em uso. Tentativa {tentativa + 1}/3")
            time.sleep(2)
    else:
        print("❌ Não foi possível remover o arquivo. Ele pode estar aberto.")

    print("✅ Processo finalizado com sucesso.")


if __name__ == "__main__":
    main()
