from datetime import datetime, timedelta
from pathlib import Path
import shutil


# --------------------------------------------------
# 📅 Calcular período padrão (caso queira reutilizar)
# --------------------------------------------------
def calcular_periodo():
    hoje = datetime.today()
    start = hoje - timedelta(days=5)
    end = hoje - timedelta(days=1)

    return start, end


# --------------------------------------------------
# 🕒 Formatar datetime no padrão do sistema
# --------------------------------------------------
def formatar_datetime_local(dt: datetime):
    return dt.strftime("%Y-%m-%dT00:00")


# --------------------------------------------------
# ♻️ Backup incremental do arquivo existente
# --------------------------------------------------
def mover_para_bkp(dest_dir: Path, bkp_dir: Path, filename: str):

    dest_file = dest_dir / filename

    if not dest_file.exists():
        return

    bkp_dir.mkdir(parents=True, exist_ok=True)

    contador = 1
    while True:
        novo_nome = f"{filename.replace('.csv','')}_{contador}.csv"
        destino_bkp = bkp_dir / novo_nome

        if not destino_bkp.exists():
            shutil.move(dest_file, destino_bkp)
            break

        contador += 1