import pandas as pd
from pathlib import Path


COLUNAS_NOVAS_3547 = [
    "ORIGEM",
    "PRODUTO",
    "EMISSOR",
    "CONTA",
    "CARTAO",
    "DT_AJUSTE",
    "DT_TRANSACAO",
    "VALOR",
    "CD_AUT",
    "MCC",
    "ARN",
    "RAZAO",
    "ONL_OFF",
    "POS",
    "NOME_EC"
]


def processar_3547(caminho_arquivo: Path):

    print(f"\n🟡 Processando 3547: {caminho_arquivo}")

    xls = pd.ExcelFile(caminho_arquivo)
    sheet_original = xls.sheet_names[0]

    print(f"📄 Sheet original: {sheet_original}")

    df = pd.read_excel(
        caminho_arquivo,
        sheet_name=sheet_original,
        dtype=str
    )

    print(f"📊 Colunas originais: {list(df.columns)}")
    print(f"📈 Linhas: {len(df)}")

    # Inserir coluna A
    df.insert(0, "ORIGEM", "")

    # Renomear colunas
    df.columns = COLUNAS_NOVAS_3547

    with pd.ExcelWriter(
        caminho_arquivo,
        engine="openpyxl",
        mode="w"
    ) as writer:
        df.to_excel(writer, sheet_name="Relatório 3547", index=False)

    print("✅ 3547 tratado com sucesso.")

    return caminho_arquivo
