from datetime import datetime, timedelta
from pathlib import Path
import time
import pandas as pd
import unicodedata
import csv


def get_days_of_month_until_yesterday(days: int = 6):
    """
    Retorna lista com os últimos N dias até ontem.
    Ex: hoje = 26/12 → retorna 20/12 a 25/12
    """
    today = datetime.today()
    end_day = today - timedelta(days=1)
    start_day = end_day - timedelta(days=days - 1)

    days_list = []
    current = start_day

    while current <= end_day:
        days_list.append(current)
        current += timedelta(days=1)

    return days_list



def format_day_range(day: datetime):
    """
    Formata: 01/12/2025 - 01/12/2025
    """
    d = day.strftime("%d/%m/%Y")
    return f"{d} - {d}"


def wait_for_download(download_dir: Path, timeout=120):
    end_time = time.time() + timeout
    while time.time() < end_time:
        files = list(download_dir.glob("*.xlsx"))
        if files:
            return max(files, key=lambda f: f.stat().st_mtime)
        time.sleep(1)
    raise TimeoutError("Download não encontrado")


def normalize_text(value):
    if not isinstance(value, str):
        return value

    # Corrige encoding quebrado
    try:
        value = value.encode("latin1").decode("utf-8")
    except Exception:
        pass

    # Remove quebras de linha internas (CRÍTICO PARA SSIS)
    value = value.replace("\r", " ")
    value = value.replace("\n", " ")
    value = value.replace("_x000D_", " ")

    # Normaliza unicode
    value = unicodedata.normalize("NFKC", value)

    # Remove múltiplos espaços gerados
    value = " ".join(value.split())

    return value


def xlsx_to_csv(xlsx_path: Path, csv_path: Path):
    df = pd.read_excel(xlsx_path, engine="openpyxl")

    for col in df.select_dtypes(include=["object"]).columns:
        df[col] = df[col].apply(normalize_text)

    csv_path.parent.mkdir(parents=True, exist_ok=True)

    df.to_csv(
        csv_path,
        index=False,
        sep=";",
        encoding="utf-8-sig",
        quotechar='"',
        quoting=csv.QUOTE_ALL,
        escapechar='\\',
        lineterminator="\n"
    )

