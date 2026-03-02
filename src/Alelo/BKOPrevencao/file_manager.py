from pathlib import Path
import shutil


def backup_arquivo(destino: Path, pasta_bkp: Path):

    if not destino.exists():
        print("ℹ️ Nenhum arquivo BASE.xlsx existente para backup.")
        return

    pasta_bkp.mkdir(parents=True, exist_ok=True)

    arquivo_bkp = pasta_bkp / "BASE_1.xlsx"

    print(f"📦 Movendo arquivo antigo para BKP: {arquivo_bkp}")

    if arquivo_bkp.exists():
        arquivo_bkp.unlink()

    shutil.move(str(destino), str(arquivo_bkp))