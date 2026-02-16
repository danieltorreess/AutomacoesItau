import pandas as pd
from pathlib import Path


# ==========================================================
# 🔎 ENCONTRAR CSV EXTRAÍDO
# ==========================================================

def encontrar_csv_extraido(pasta_downloads: Path):
    """
    Procura recursivamente pelo CSV de Desk Messages
    dentro da pasta de Downloads.
    """

    arquivos = list(
        pasta_downloads.rglob(
            "atendhumanoalelopodprd_Desk_Messages_*.csv"
        )
    )

    if not arquivos:
        return None

    # Ordena pelo mais recente
    arquivos.sort(key=lambda x: x.stat().st_mtime, reverse=True)

    return arquivos[0]


# ==========================================================
# 🔎 LEITURA ROBUSTA DE CSV
# ==========================================================

def ler_csv_robusto(caminho):
    """
    Tenta ler CSV com diferentes separadores e encodings.
    Retorna também o separador usado.
    """

    tentativas = [
        {"sep": ";", "encoding": "utf-8-sig"},
        {"sep": ";", "encoding": "latin1"},
        {"sep": ",", "encoding": "utf-8-sig"},
        {"sep": ",", "encoding": "latin1"},
    ]

    for config in tentativas:
        try:
            df = pd.read_csv(
                caminho,
                sep=config["sep"],
                encoding=config["encoding"],
                low_memory=False
            )

            print(f"✔ CSV lido com sep='{config['sep']}' encoding='{config['encoding']}'")

            return df, config["sep"]

        except Exception:
            continue

    raise Exception(f"❌ Não foi possível ler o CSV: {caminho}")


# ==========================================================
# 📅 IDENTIFICAR ÚLTIMA DATA DO CONSOLIDADO
# ==========================================================

def obter_ultima_data_rede(caminho_csv_rede: Path):

    df_rede, separador = ler_csv_robusto(caminho_csv_rede)

    if df_rede.empty:
        raise Exception("❌ O CSV da rede está vazio.")

    if df_rede.shape[1] < 7:
        raise Exception("❌ O CSV da rede não possui coluna G suficiente.")

    coluna_data = df_rede.iloc[:, 6]

    coluna_data = pd.to_datetime(
        coluna_data,
        dayfirst=True,
        errors="coerce"
    )

    ultima_data = coluna_data.max()

    if pd.isna(ultima_data):
        raise Exception("❌ Não foi possível identificar a última data válida na coluna G.")

    print(f"📅 Última data encontrada na rede: {ultima_data.date()}")

    return ultima_data, df_rede, separador


# ==========================================================
# 🔎 FILTRAR REGISTROS NOVOS
# ==========================================================

def filtrar_novos_registros(caminho_csv_novo: Path, ultima_data):

    df_novo, _ = ler_csv_robusto(caminho_csv_novo)

    if df_novo.empty:
        print("⚠️ CSV novo está vazio.")
        return pd.DataFrame()

    if df_novo.shape[1] < 7:
        raise Exception("❌ CSV novo não possui coluna G suficiente.")

    coluna_data = df_novo.iloc[:, 6]

    coluna_data = pd.to_datetime(
        coluna_data,
        dayfirst=True,
        errors="coerce"
    )

    df_novo["__DATA_CONVERTIDA__"] = coluna_data

    df_filtrado = df_novo[
        df_novo["__DATA_CONVERTIDA__"] > ultima_data
    ].copy()

    df_filtrado.drop(columns=["__DATA_CONVERTIDA__"], inplace=True)

    print(f"🔎 Registros após filtro incremental: {len(df_filtrado)}")

    return df_filtrado


# ==========================================================
# ➕ APPEND NO CONSOLIDADO
# ==========================================================

def append_no_csv_rede(df_rede, df_novos, caminho_csv_rede: Path, separador):

    if df_novos.empty:
        print("⚠️ Nenhum novo registro para inserir.")
        return

    # Escreve SOMENTE os novos registros no final do arquivo
    df_novos.to_csv(
        caminho_csv_rede,
        mode="a",              # append
        sep=separador,
        index=False,
        header=False           # NÃO escrever cabeçalho novamente
    )

    print(f"✅ {len(df_novos)} registros inseridos com sucesso (append direto).")
