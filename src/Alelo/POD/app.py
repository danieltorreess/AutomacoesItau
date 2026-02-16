from pathlib import Path
import shutil
import pandas as pd
from datetime import datetime

from src.Alelo.POD.file_finder import encontrar_zip_mais_recente
from src.Alelo.POD.zip_handler import extrair_zip
from src.Alelo.POD.csv_processor import encontrar_csv_extraido, ler_csv_robusto


# ==========================================================
# 📌 CONFIGURAÇÕES
# ==========================================================

PASTA_DOWNLOADS = Path(r"C:\Users\ab1541240\Downloads")

CAMINHO_BASE = Path(r"\\BRSBESRV960\Publico\BASE\ALELO\POD")

ARQUIVO_ATIVO = CAMINHO_BASE / "atendhumanoalelopodprd_Desk_Messages.csv"

PASTA_BKP = CAMINHO_BASE / "BKP"


# ==========================================================
# 📅 IDENTIFICAR ÚLTIMA DATA DA COLUNA G
# ==========================================================
def obter_ultima_data_coluna_g(caminho_csv):

    df, _ = ler_csv_robusto(caminho_csv)

    if df.shape[1] < 7:
        raise Exception("CSV não possui coluna G suficiente.")

    datas_validas = []

    for valor in df.iloc[:, 6]:

        if pd.isna(valor):
            continue

        valor_str = str(valor).strip()

        # Remove horário se existir
        valor_str = valor_str.split(" ")[0]

        formatos_possiveis = [
            "%d/%m/%Y",
            "%Y-%m-%d",
            "%d-%m-%Y"
        ]

        for formato in formatos_possiveis:
            try:
                data_convertida = datetime.strptime(valor_str, formato)
                datas_validas.append(data_convertida)
                break
            except ValueError:
                continue

    if not datas_validas:
        # Debug para entender o formato real
        exemplos = df.iloc[:5, 6].tolist()
        raise Exception(
            f"Não foi possível identificar datas válidas. Exemplos encontrados na coluna G: {exemplos}"
        )

    ultima_data = max(datas_validas)

    return ultima_data


# ==========================================================
# 🚀 MAIN
# ==========================================================

def main():

    print("\n🚀 Iniciando RPA POD - Nova Arquitetura")

    # ------------------------------------------------------
    # 1️⃣ Se existir arquivo ativo → renomear e mover p/ BKP
    # ------------------------------------------------------

    if ARQUIVO_ATIVO.exists():

        print("📂 Arquivo ativo encontrado.")

        ultima_data = obter_ultima_data_coluna_g(ARQUIVO_ATIVO)

        print(f"📅 Última data identificada: {ultima_data}")

        data_formatada = ultima_data.strftime("%Y-%m-%d")

        novo_nome_bkp = f"atendhumanoalelopodprd_Desk_Messages-{data_formatada}.csv"

        destino_bkp = PASTA_BKP / novo_nome_bkp

        PASTA_BKP.mkdir(exist_ok=True)

        shutil.move(str(ARQUIVO_ATIVO), destino_bkp)

        print(f"📦 Arquivo movido para BKP: {destino_bkp.name}")

    else:
        print("ℹ️ Nenhum arquivo ativo encontrado (primeira execução).")

    # ------------------------------------------------------
    # 2️⃣ Encontrar ZIP mais recente
    # ------------------------------------------------------

    zip_file = encontrar_zip_mais_recente(PASTA_DOWNLOADS)

    if not zip_file:
        print("❌ Nenhum ZIP encontrado.")
        return

    print(f"📦 ZIP encontrado: {zip_file.name}")

    # ------------------------------------------------------
    # 3️⃣ Extrair ZIP
    # ------------------------------------------------------

    pasta_extraida = extrair_zip(zip_file, PASTA_DOWNLOADS)

    print(f"📂 Extraído para: {pasta_extraida}")

    # ------------------------------------------------------
    # 4️⃣ Encontrar CSV extraído
    # ------------------------------------------------------

    csv_extraido = encontrar_csv_extraido(PASTA_DOWNLOADS)

    if not csv_extraido:
        print("❌ CSV extraído não encontrado.")
        return

    print(f"📄 CSV encontrado: {csv_extraido.name}")

    # ------------------------------------------------------
    # 5️⃣ Renomear e mover para pasta principal
    # ------------------------------------------------------

    destino_final = CAMINHO_BASE / "atendhumanoalelopodprd_Desk_Messages.csv"

    shutil.move(str(csv_extraido), destino_final)

    print(f"✅ Novo arquivo salvo como: {destino_final.name}")

    print("🏁 Processo finalizado com sucesso.")


# ==========================================================

if __name__ == "__main__":
    main()
