from pathlib import Path

from .config import BASE_PATH
from .csv_transformer import corrigir_csv


def main():

    print("\n🚀 Iniciando correção de CSV\n")

    arquivos = list(BASE_PATH.glob("*.csv"))

    print(f"📂 {len(arquivos)} arquivos encontrados\n")

    for arquivo in arquivos:

        corrigir_csv(arquivo)

    print("\n🏁 Processo finalizado\n")


if __name__ == "__main__":
    main()