import shutil
from pathlib import Path
import pandas as pd


PASTA_REDE = Path(r"\\brsbesrv960\publico\REPORTS\ALELO\MOTIVO_CHAMADOR_POD")
PASTA_BKP = PASTA_REDE / "BKP"

NOME_FINAL = "Motivo Chamador - POD.csv"


def encontrar_arquivo_xlsx():
    arquivos = list(PASTA_REDE.glob("Motivo Chamador - POD*.xlsx"))

    if not arquivos:
        print("❌ Nenhum arquivo encontrado.")
        return None

    # pega o mais recente
    arquivo = max(arquivos, key=lambda f: f.stat().st_mtime)

    print(f"📄 Arquivo encontrado: {arquivo.name}")
    return arquivo


def limpar_dataframe(df: pd.DataFrame) -> pd.DataFrame:

    print("🧹 Removendo ';' e quebras de linha internas...")

    df = df.astype(str)

    df = df.replace(";", "", regex=True)
    df = df.replace(r"\r", "", regex=True)
    df = df.replace(r"\n", "", regex=True)

    return df


def gerar_bkp_incremental():
    PASTA_BKP.mkdir(parents=True, exist_ok=True)

    contador = 1

    while True:
        nome_bkp = f"Motivo Chamador - POD_{contador}.csv"
        caminho_bkp = PASTA_BKP / nome_bkp

        if not caminho_bkp.exists():
            return caminho_bkp

        contador += 1


def main():

    print("\n🚀 Iniciando RPA Motivo Chamador - POD")
    print("===================================================")

    arquivo_xlsx = encontrar_arquivo_xlsx()

    if not arquivo_xlsx:
        return

    print("📖 Lendo Excel...")
    df = pd.read_excel(arquivo_xlsx, dtype=str)

    print(f"📊 Linhas: {len(df)}")
    print(f"📊 Colunas: {len(df.columns)}")

    df = limpar_dataframe(df)

    caminho_csv_final = PASTA_REDE / NOME_FINAL

    # 🔥 BKP se já existir
    if caminho_csv_final.exists():
        caminho_bkp = gerar_bkp_incremental()
        print(f"📦 Movendo arquivo antigo para BKP: {caminho_bkp.name}")
        shutil.move(caminho_csv_final, caminho_bkp)

    print("💾 Salvando novo CSV...")

    df.to_csv(
        caminho_csv_final,
        index=False,
        encoding="utf-8-sig",
        sep=";"
    )


    print("✅ Processo finalizado com sucesso.")


if __name__ == "__main__":
    main()
