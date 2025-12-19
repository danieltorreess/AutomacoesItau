from datetime import datetime
import time
import shutil
from pathlib import Path

def get_date_range():
    hoje = datetime.today()
    inicio = hoje.replace(day=1).strftime("%d/%m/%Y")
    fim = hoje.strftime("%d/%m/%Y")
    return f"{inicio} - {fim}"

def wait_for_download(download_dir: Path, timeout=120):
    end_time = time.time() + timeout
    while time.time() < end_time:
        files = list(download_dir.glob("download_*.xlsx"))
        if files:
            return max(files, key=lambda f: f.stat().st_mtime)
        time.sleep(1)
    raise TimeoutError("Download não encontrado")

def move_and_rename(file: Path, dest_dir: Path, new_name: str):
    dest_dir.mkdir(parents=True, exist_ok=True)
    target = dest_dir / new_name
    shutil.move(str(file), str(target))
