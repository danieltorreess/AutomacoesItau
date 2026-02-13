from pathlib import Path
import shutil

from src.Alelo.BeedooLocal.processor import (
    encontrar_arquivo,
    processar_excel_para_csv
)
from src.Alelo.BeedooLocal.file_utils import mover_para_bkp


# ==========================
# CONFIG
# ==========================

DOWNLOAD_DIR = Path.home() / "Downloads"

DESTINO_REDE = Path(r"\\brsbesrv960\publico\REPORTS\ALELO\BEEDOO")

BKP_POST = DESTINO_REDE / "BKP" / "post-analitic"
BKP_VIEWS = DESTINO_REDE / "BKP" / "views-track-detailed"

DESTINO_REDE.mkdir(parents=True, exist_ok=True)


def main():

    print("🚀 Iniciando RPA BEEDOO LOCAL")

    # ==============================
    # 1️⃣ POST ANALITIC
    # ==============================

    arquivo_post = encontrar_arquivo("post-analitic")

    if arquivo_post:
        print(f"📥 Encontrado: {arquivo_post.name}")

        destino_csv = DESTINO_REDE / "post-analitic.csv"

        mover_para_bkp(destino_csv, BKP_POST)

        processar_excel_para_csv(
            arquivo_post,
            "post-analitic",
            DESTINO_REDE
        )

        arquivo_post.unlink()  # remove xlsx dos downloads

    else:
        print("⚠️ Arquivo post-analitic não encontrado.")

    # ==============================
    # 2️⃣ VIEWS TRACK DETAILED
    # ==============================

    arquivo_views = encontrar_arquivo("views-track-detailed")

    if arquivo_views:
        print(f"📥 Encontrado: {arquivo_views.name}")

        destino_csv = DESTINO_REDE / "views-track-detailed.csv"

        mover_para_bkp(destino_csv, BKP_VIEWS)

        processar_excel_para_csv(
            arquivo_views,
            "views-track-detailed",
            DESTINO_REDE
        )

        arquivo_views.unlink()

    else:
        print("⚠️ Arquivo views-track-detailed não encontrado.")

    print("✅ Processo finalizado com sucesso.")


if __name__ == "__main__":
    main()
