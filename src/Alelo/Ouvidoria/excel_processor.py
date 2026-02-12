import pandas as pd


SHEETS_VALIDAS = ["BASE ATENTO", "FINALIZADOS"]


def processar_arquivo_excel(caminho_arquivo):
    """
    - Mantém apenas as sheets desejadas
    - Mantém apenas colunas A até P
    - Converte coluna C para texto
    """

    print("🛠 Processando arquivo Excel...")

    xls = pd.ExcelFile(caminho_arquivo)

    sheets_encontradas = []
    dados_processados = {}

    # 🔍 Verifica quais sheets válidas existem
    for sheet in xls.sheet_names:
        nome_normalizado = sheet.strip().upper()

        if nome_normalizado in SHEETS_VALIDAS:
            sheets_encontradas.append(sheet)

    if not sheets_encontradas:
        raise Exception(
            f"Nenhuma das sheets esperadas foi encontrada. "
            f"Sheets disponíveis: {xls.sheet_names}"
        )

    # 🔄 Processa apenas as sheets válidas encontradas
    for sheet in sheets_encontradas:
        df = pd.read_excel(
            caminho_arquivo,
            sheet_name=sheet,
            dtype=str  # 🔥 força tudo como string (evita perda de zeros)
        )

        # Mantém apenas colunas A até P (0 até 15)
        df = df.iloc[:, :16]

        # Garante coluna C como string
        if df.shape[1] >= 3:
            df.iloc[:, 2] = df.iloc[:, 2].astype(str)

        dados_processados[sheet] = df

    # 🔥 Só agora sobrescreve o arquivo
    with pd.ExcelWriter(
        caminho_arquivo,
        engine="openpyxl",
        mode="w"
    ) as writer:

        for sheet, df in dados_processados.items():
            df.to_excel(writer, sheet_name=sheet, index=False)

    print("✅ Tratamento concluído.")
