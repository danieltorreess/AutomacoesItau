import pandas as pd
from pathlib import Path


COLUNAS_NOVAS = [
    "NOME_ORIGEM",
    "BASE_ORIGEM",
    "PRODUTO",
    "DATA",
    "TIPO_AJUSTE",
    "CARTAO",
    "CONTA",
    "CREDITO",
    "DEBITO",
    "LOGIN",
    "NOME",
    "PERFIL"
]


def processar_3580(caminho_arquivo: Path):

    print(f"\n🟢 Processando 3580: {caminho_arquivo}")

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

    df.insert(0, "NOME_ORIGEM", "")
    df.columns = COLUNAS_NOVAS

    with pd.ExcelWriter(
        caminho_arquivo,
        engine="openpyxl",
        mode="w"
    ) as writer:
        df.to_excel(writer, sheet_name="3580", index=False)

    print("✅ 3580 tratado com sucesso.")

    return caminho_arquivo
