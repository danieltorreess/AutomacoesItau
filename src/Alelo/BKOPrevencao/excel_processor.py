import pandas as pd
from pathlib import Path


COLUNAS_ESPERADAS = [
    "Referencia",
    "DT Abertura",
    "Prz Sol",
    "Dt Sol",
    "Evento",
    "Sub Evento",
    "Status",
    "Area Dep",
    "VIP",
    "DT Follow",
    "Operador",
    "Cartão"
]


import pandas as pd
from datetime import datetime


def converter_data_excel(valor):

    if pd.isna(valor) or valor == "":
        return ""

    valor = str(valor).strip()

    # Se for número serial Excel
    if valor.replace(".", "", 1).isdigit():
        try:
            data = pd.to_datetime("1899-12-30") + pd.to_timedelta(float(valor), "D")
            return data.strftime("%d/%m/%Y %H:%M:%S")
        except:
            return ""

    # Se for string data normal
    try:
        data = pd.to_datetime(valor, dayfirst=True, errors="coerce")
        if pd.isna(data):
            return ""
        return data.strftime("%d/%m/%Y %H:%M:%S")
    except:
        return ""



def formatar_datas(df):

    for coluna in ["DT Abertura", "Prz Sol", "Dt Sol"]:

        if coluna in df.columns:
            df[coluna] = df[coluna].apply(converter_data_excel)

    return df


def ajustar_colunas(df):

    # Remove espaços invisíveis
    df.columns = df.columns.str.strip()

    # Mantém apenas as colunas esperadas que existirem
    colunas_presentes = [c for c in COLUNAS_ESPERADAS if c in df.columns]

    df = df[colunas_presentes].copy()

    # Garante coluna Status extra no final
    # df["Status"] = df["Status"] if "Status" in df.columns else ""
    # df["Status_Extra"] = ""

    # df = df.rename(columns={"Status_Extra": "Status"})

    return df

def tratar_cartao(df):

    if "Cartão" not in df.columns:
        return df

    # Remove espaços
    df["Cartão"] = df["Cartão"].str.strip()

    # Remove .0 caso exista
    df["Cartão"] = df["Cartão"].str.replace(".0", "", regex=False)

    # Corrige notação científica se vier como string tipo 5.06758E+15
    def corrigir_notacao(valor):
        try:
            if "E+" in valor.upper():
                return str(int(float(valor)))
            return valor
        except:
            return valor

    df["Cartão"] = df["Cartão"].apply(corrigir_notacao)

    return df

def processar_excel(caminho_arquivo: Path, destino_final: Path):

    print("📊 Lendo arquivo Excel...")
    with pd.ExcelFile(caminho_arquivo) as xls:

        with pd.ExcelWriter(destino_final, engine="openpyxl") as writer:

            sheets_processadas = 0

            for sheet in xls.sheet_names:

                df = pd.read_excel(
                    xls,
                    sheet_name=sheet,
                    dtype=str
                )

                if sheet.strip().lower() == "prevenção":
                    novo_nome = "PREVENÇÃO"

                elif "intercâmbio" in sheet.lower():
                    novo_nome = "INTERCÂMBIO"

                else:
                    continue

                print(f"🛠 Processando sheet: {sheet} -> {novo_nome}")
                print(f"🔎 Colunas originais: {list(df.columns)}")

                df = ajustar_colunas(df)
                df = formatar_datas(df)
                df = tratar_cartao(df)

                print(f"✅ Total de linhas: {len(df)}")

                df.to_excel(writer, sheet_name=novo_nome, index=False)

                sheets_processadas += 1

            if sheets_processadas == 0:
                raise Exception("Nenhuma sheet válida foi encontrada no arquivo.")

    print("✅ Excel processado com sucesso.")
    
