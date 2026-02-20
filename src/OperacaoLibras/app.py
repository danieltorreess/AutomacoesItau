from pathlib import Path
import shutil
from openpyxl import load_workbook
from datetime import datetime
import win32com.client as win32

# ================= CONFIG =================

DOWNLOADS = Path.home() / "Downloads"

PASTA_BASE = Path(r"\\BRSBESRV960\Publico\REPORTS\ITAU\JOURNEY_LIBRAS\Pesquisa")
PASTA_BKP = PASTA_BASE / "BKP"

PADRAO_ARQUIVO = "App - Pesquisas de Atendimento Atualizado*.xlsx"

NOME_FIXO_REDE = "PesquisaAtendimento.xlsx"
NOME_BASE_BKP = "PesquisaAtendimento"

# ==========================================


def localizar_arquivo_recente():
    arquivos = list(DOWNLOADS.glob(PADRAO_ARQUIVO))

    if not arquivos:
        return None

    # pega o mais recente
    return max(arquivos, key=lambda f: f.stat().st_mtime)

def remover_filtros_excel(caminho_arquivo: Path):
    print("🧹 Removendo filtros via Excel COM...")

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    wb = excel.Workbooks.Open(str(caminho_arquivo))

    try:
        ws = wb.Worksheets("Exportacao")

        # Remove AutoFilter
        if ws.AutoFilterMode:
            ws.AutoFilterMode = False

        # Caso seja tabela estruturada
        for tabela in ws.ListObjects:
            if tabela.ShowAutoFilter:
                tabela.ShowAutoFilter = False

        wb.Save()
        print("✅ Filtros removidos com sucesso.")

    finally:
        wb.Close(SaveChanges=True)
        excel.Quit()


def gerar_nome_bkp_incremental():
    contador = 1

    while True:
        nome = f"{NOME_BASE_BKP}_{contador}.xlsx"
        caminho = PASTA_BKP / nome

        if not caminho.exists():
            return caminho

        contador += 1


def mover_para_bkp():
    arquivo_atual = PASTA_BASE / NOME_FIXO_REDE

    if not arquivo_atual.exists():
        return

    PASTA_BKP.mkdir(exist_ok=True)

    novo_nome = gerar_nome_bkp_incremental()

    shutil.move(arquivo_atual, novo_nome)

    print(f"📦 Backup criado: {novo_nome}")


def mover_para_rede(origem: Path):
    destino = PASTA_BASE / NOME_FIXO_REDE
    shutil.move(origem, destino)
    print(f"🚚 Arquivo movido para rede: {destino}")


def main():
    print("\n🚀 Iniciando RPA - Pesquisa Atendimento (Modo Downloads)")

    # 1️⃣ Localiza arquivo no Downloads
    arquivo = localizar_arquivo_recente()

    if not arquivo:
        print("❌ Nenhum arquivo encontrado nos Downloads.")
        return

    print(f"📂 Arquivo encontrado: {arquivo.name}")

    # 2️⃣ Remove filtros
    remover_filtros_excel(arquivo)

    # 3️⃣ Move arquivo atual da rede para BKP
    mover_para_bkp()

    # 4️⃣ Move novo arquivo para rede com nome fixo
    mover_para_rede(arquivo)

    print("🎯 Processo finalizado com sucesso.")
    print("📊 Arquivo pronto para SSIS")


if __name__ == "__main__":
    main()