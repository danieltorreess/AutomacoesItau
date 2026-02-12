from pathlib import Path


def contar_linhas_csv(caminho_arquivo: Path) -> int:
    """
    Conta linhas de um arquivo CSV sem carregar tudo na memória.
    Performance: O(n) em tempo e O(1) em memória.
    """
    with caminho_arquivo.open("r", encoding="utf-8", errors="ignore") as f:
        return sum(1 for _ in f)


def main():
    # Pega a raiz do projeto automaticamente
    raiz_projeto = Path(__file__).resolve().parents[2]

    arquivo_csv = raiz_projeto / "UserItemStatusCSV_V2 (1).csv"

    if not arquivo_csv.exists():
        print(f"Arquivo não encontrado: {arquivo_csv}")
        return

    print(f"Lendo arquivo: {arquivo_csv}")
    total_linhas = contar_linhas_csv(arquivo_csv)

    print(f"\nTotal de linhas: {total_linhas:,}")


if __name__ == "__main__":
    main()
