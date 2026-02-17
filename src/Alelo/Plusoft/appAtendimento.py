##type:ignore
import shutil
from pathlib import Path
import pandas as pd


PASTA_REDE = Path(
    r"\\brsbesrv960\Publico\REPORTS\001 CARGAS\001 EXCEL\ALELO CAU\EMAIL ATENDIMENTO"
)


def encontrar_arquivo(padrao):
    arquivos = list(PASTA_REDE.glob(padrao))

    if not arquivos:
        print(f"❌ Nenhum arquivo encontrado para padrão: {padrao}")
        return None

    arquivo = max(arquivos, key=lambda f: f.stat().st_mtime)
    print(f"📄 Arquivo encontrado: {arquivo.name}")

    return arquivo


def limpar_dataframe(df: pd.DataFrame):

    print("🧹 Removendo ';' e quebras de linha internas...")

    df = df.astype(str)

    df = df.applymap(lambda x: str(x).replace(";", ""))
    df = df.replace(r"\r", "", regex=True)
    df = df.replace(r"\n", "", regex=True)

    return df


def gerar_bkp_incremental(pasta_bkp: Path, nome_base: str):

    pasta_bkp.mkdir(parents=True, exist_ok=True)

    contador = 1

    while True:
        nome_bkp = f"{nome_base}_{contador}.csv"
        caminho_bkp = pasta_bkp / nome_bkp

        if not caminho_bkp.exists():
            return caminho_bkp

        contador += 1


def processar_arquivo(padrao, nome_sheet, nome_final, subpasta_bkp):

    print("\n---------------------------------------------------")

    arquivo_xlsx = encontrar_arquivo(padrao)

    if not arquivo_xlsx:
        return

    print("📖 Lendo Excel...")
    df = pd.read_excel(arquivo_xlsx, dtype=str)

    print(f"📊 Linhas: {len(df)}")
    print(f"📊 Colunas: {len(df.columns)}")

    # 🔥 Renomeia sheet (regravando Excel antes do CSV)
    with pd.ExcelWriter(
        arquivo_xlsx,
        engine="openpyxl",
        mode="w"
    ) as writer:
        df.to_excel(writer, sheet_name=nome_sheet, index=False)

    df = limpar_dataframe(df)

    caminho_csv_final = PASTA_REDE / nome_final
    pasta_bkp = PASTA_REDE / "BKP" / subpasta_bkp

    # 🔥 Backup se já existir
    if caminho_csv_final.exists():
        caminho_bkp = gerar_bkp_incremental(pasta_bkp, nome_final.replace(".csv", ""))
        print(f"📦 Movendo arquivo antigo para BKP: {caminho_bkp.name}")
        shutil.move(caminho_csv_final, caminho_bkp)

    print("💾 Salvando novo CSV...")

    df.to_csv(
        caminho_csv_final,
        index=False,
        encoding="utf-8-sig",
        sep=";"
    )

    print("✅ Arquivo processado com sucesso.")


def main():

    print("\n🚀 Iniciando RPA Plusoft - Atendimento")
    print("===================================================")

    # 🔹 MAY
    processar_arquivo(
        padrao="Email POD CAE - MAY*.xlsx",
        nome_sheet="Email POD CAE - MAY",
        nome_final="Email POD CAE - MAY.csv",
        subpasta_bkp="May"
    )

    # 🔹 ELIZA
    processar_arquivo(
        padrao="Franquia Eliza*.xlsx",
        nome_sheet="Franquia Eliza",
        nome_final="Franquia Eliza.csv",
        subpasta_bkp="Eliza"
    )

    print("\n🏁 Processo finalizado.")


if __name__ == "__main__":
    main()
