import pandas as pd
from pathlib import Path


DOWNLOAD_DIR = Path.home() / "Downloads"


def encontrar_planilha():

    arquivos = list(DOWNLOAD_DIR.glob("Planilha de atribuição*.xlsx"))

    if not arquivos:
        return None

    arquivos.sort(key=lambda x: x.stat().st_mtime, reverse=True)

    return arquivos[0]


def processar_planilha(caminho_origem: Path, caminho_destino: Path):

    print(f"🔄 Processando {caminho_origem.name}")

    # ==============================
    # 🔹 Abre Excel garantindo fechamento correto
    # ==============================
    with pd.ExcelFile(caminho_origem) as xls:

        sheet_alvo = None

        for sheet in xls.sheet_names:
            if sheet.startswith("Atribuições"):
                sheet_alvo = sheet
                break

        if not sheet_alvo:
            raise Exception("❌ Nenhuma sheet 'Atribuições*' encontrada.")

        df = pd.read_excel(xls, sheet_name=sheet_alvo, dtype=str)

    # ==============================
    # 🔹 Salva apenas a sheet necessária
    # ==============================
    with pd.ExcelWriter(caminho_destino, engine="openpyxl", mode="w") as writer:
        df.to_excel(writer, sheet_name="Atribuicoes", index=False)

    print(f"✅ Arquivo padronizado salvo: {caminho_destino}")
