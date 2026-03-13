from pathlib import Path
import pandas as pd
import shutil

from .config import (
    DOWNLOADS,
    PATTERN,
    NETWORK_PATH,
    BKP_PATH,
    FINAL_FILENAME,
    SHEET_NAME
)


def get_latest_download() -> Path:

    files = list(DOWNLOADS.glob(PATTERN))

    if not files:
        raise FileNotFoundError(
            "Nenhum arquivo callflex_agente_detalhado encontrado em Downloads"
        )

    latest_file = max(files, key=lambda f: f.stat().st_mtime)

    return latest_file


def read_excel(file_path: Path) -> pd.DataFrame:

    tables = pd.read_html(file_path)

    if not tables:
        raise Exception("Nenhuma tabela encontrada no relatório")

    df = tables[0]

    # header correto
    df.columns = df.iloc[0]
    df = df[1:]

    # limpar colunas
    df.columns = df.columns.astype(str).str.strip()

    # remover linhas vazias
    df = df.dropna(how="all")

    # reset index
    df = df.reset_index(drop=True)

    return df

def get_max_data_login(df: pd.DataFrame) -> str:

    datas = pd.to_datetime(df["DATA_LOGIN"], errors="coerce")

    max_date = datas.max()

    return max_date.strftime("%Y-%m-%d")


def backup_existing_file(max_date: str):

    current_file = NETWORK_PATH / FINAL_FILENAME

    if not current_file.exists():
        return

    BKP_PATH.mkdir(parents=True, exist_ok=True)

    backup_name = f"PerformanceAgentes_{max_date}.xlsx"

    backup_path = BKP_PATH / backup_name

    shutil.move(current_file, backup_path)

    print(f"Backup criado: {backup_path}")


def save_new_file(df: pd.DataFrame):

    output_path = NETWORK_PATH / FINAL_FILENAME

    df.to_excel(
        output_path,
        sheet_name=SHEET_NAME,
        index=False,
        engine="openpyxl"
    )

    print(f"Arquivo salvo: {output_path}")


def run_pipeline():

    print("Iniciando RPA BaseTempos")

    source_file = get_latest_download()
    print(f"Arquivo encontrado: {source_file}")

    df = read_excel(source_file)

    max_date = get_max_data_login(df)
    print(f"Maior DATA_LOGIN: {max_date}")

    backup_existing_file(max_date)

    save_new_file(df)

    print("RPA finalizado com sucesso")