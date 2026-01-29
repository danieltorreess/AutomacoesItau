import win32com.client as win32
import pandas as pd
import time
from pathlib import Path


BASE_PATH = Path(
    r"C:\Users\ab1541240\OneDrive - ATENTO Global\Documents\ITAU\RelatoriosItau\02 - PowerBI\MIS00000 - SMS_EMAILS"
)

ARQUIVO_ORIGEM = BASE_PATH / "Login e Logout LD_SEM_PROTECAO.xlsx"
ARQUIVO_SAIDA = BASE_PATH / "LoginLogoutOficial.xlsx"

ABA_TABELA = "Tempos_analistas"


def extrair_base_completa():
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    wb = excel.Workbooks.Open(str(ARQUIVO_ORIGEM))
    ws = wb.Worksheets(ABA_TABELA)

    used = ws.UsedRange

    # 🔎 achar linha do cabeçalho "Analistas" (coluna B)
    linha_header = None
    linha_total = None

    for r in range(1, used.Rows.Count + 1):
        v = ws.Cells(r, 2).Value
        if isinstance(v, str) and v.strip().lower() == "analistas":
            linha_header = r
            break

    if not linha_header:
        raise RuntimeError("Não encontrei 'Analistas'")

    for r in range(linha_header + 1, used.Rows.Count + 1):
        v = ws.Cells(r, 2).Value
        if isinstance(v, str) and v.strip().lower() == "total geral":
            linha_total = r
            break

    primeira = linha_header + 1
    ultima = linha_total - 1

    print(f"📌 Analistas encontrados da linha {primeira} até {ultima}")

    dfs = []
    cabecalho_final = None

    for r in range(primeira, ultima + 1):
        analista = ws.Cells(r, 2).Value
        if not analista:
            continue

        print(f"▶ Explodindo analista: {analista}")

        ws.Cells(r, 3).ShowDetail = True
        time.sleep(0.3)

        aba = wb.ActiveSheet

        # 🔎 achar a PRIMEIRA linha com valor na coluna A
# 🔎 achar a linha de CABEÇALHO REAL (colunas db_)
        header_row = None
        for i in range(1, aba.UsedRange.Rows.Count + 1):
            v = aba.Cells(i, 1).Value
            if isinstance(v, str) and v.strip().lower().startswith("db_"):
                header_row = i
                break

        if not header_row:
            aba.Delete()
            continue

        data_start = header_row + 1

        last_row = aba.Cells(aba.Rows.Count, 1).End(-4162).Row  # xlUp
        last_col = aba.Cells(header_row, aba.Columns.Count).End(-4159).Column  # xlToLeft

        # 📋 ler cabeçalho
        headers = [
            str(aba.Cells(header_row, c).Value).strip()
            for c in range(1, last_col + 1)
        ]

        # 📋 ler dados
        dados = []
        for i in range(data_start, last_row + 1):
            row = [
                "" if aba.Cells(i, c).Value is None else str(aba.Cells(i, c).Value)
                for c in range(1, last_col + 1)
            ]
            dados.append(row)

        df = pd.DataFrame(dados, columns=headers)

        if cabecalho_final is None:
            cabecalho_final = headers
        else:
            df = df[cabecalho_final]

        dfs.append(df)

        aba.Delete()

    wb.Close(False)
    excel.Quit()

    if not dfs:
        raise RuntimeError("Nenhum dado extraído")

    df_final = pd.concat(dfs, ignore_index=True)
    df_final = df_final.astype(str)

    df_final.to_excel(ARQUIVO_SAIDA, index=False)

    print(f"\n✅ Arquivo gerado com sucesso:\n{ARQUIVO_SAIDA}")


if __name__ == "__main__":
    extrair_base_completa()
