from datetime import datetime, timedelta
import time
import shutil
from pathlib import Path
import pandas as pd
import unicodedata
import csv

def get_date_range(days: int = 6):
    hoje = datetime.today()
    inicio = hoje - timedelta(days=days)
    return f"{inicio.strftime('%d/%m/%Y')} - {hoje.strftime('%d/%m/%Y')}"

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

def xlsx_to_csv(xlsx_path: Path, csv_path: Path):
    """
    Converte um arquivo XLSX para CSV (UTF-8, ; como separador).
    """
    df = pd.read_excel(xlsx_path, engine="openpyxl")

    # Normaliza todas as colunas do tipo texto
    for col in df.select_dtypes(include=["object"]).columns:
        df[col] = df[col].apply(normalize_text)

    csv_path.parent.mkdir(parents=True, exist_ok=True)

    df.to_csv(
        csv_path,
        index=False,
        sep=";",          # padrão amigável para SSIS / Excel PT-BR
        encoding="utf-8-sig",
        quotechar='"',
        quoting=csv.QUOTE_ALL,
        escapechar='\\',
        lineterminator="\n"
    )

def normalize_text(value):
    if not isinstance(value, str):
        return value

    # Corrige encoding quebrado (ex: NÃ£o → Não)
    try:
        value = value.encode("latin1").decode("utf-8")
    except Exception:
        pass

    # Normalização unicode
    value = unicodedata.normalize("NFKC", value)

    # 🔥 REMOVE QUEBRAS DE LINHA INTERNAS
    value = value.replace("\r\n", " ")
    value = value.replace("\n", " ")
    value = value.replace("\r", " ")

    # 🔥 REMOVE MARCADORES DO EXCEL / XML
    value = value.replace("_x000D_", " ")
    value = value.replace("_x000A_", " ")

    # Limpeza final
    value = " ".join(value.split())

    return value
