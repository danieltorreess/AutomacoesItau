from datetime import datetime, timedelta
from pathlib import Path
import time
import pandas as pd
import unicodedata
import csv


def get_days_of_month_until_yesterday():
    """
    Retorna lista de datas do dia 01 até ontem.
    """
    today = datetime.today()
    first_day = today.replace(day=1)
    last_day = today - timedelta(days=1)

    days = []
    current = first_day

    while current <= last_day:
        days.append(current)
        current += timedelta(days=1)

    return days


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
    try:
        value = value.encode("latin1").decode("utf-8")
    except Exception:
        pass
    return unicodedata.normalize("NFKC", value)


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
