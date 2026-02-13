import pandas as pd
from pathlib import Path


NOME_SHEET_PADRAO = "Status das Demandas | Operações"


def processar_excel(caminho_origem: Path, caminho_destino: Path):

    print(f"🔄 Processando {caminho_origem.name}")

    with pd.ExcelFile(caminho_origem) as xls:

        # Pega a primeira sheet (assumindo que é a principal)
        sheet_original = xls.sheet_names[0]

        df = pd.read_excel(xls, sheet_name=sheet_original, dtype=str)

    # Salva com nome padronizado da sheet
    with pd.ExcelWriter(caminho_destino, engine="openpyxl", mode="w") as writer:
        df.to_excel(writer, sheet_name=NOME_SHEET_PADRAO, index=False)

    print(f"✅ Arquivo padronizado salvo: {caminho_destino}")
