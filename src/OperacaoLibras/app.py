from pathlib import Path
import shutil
import win32com.client as win32
import logging
from datetime import datetime
import time
import sys

# ================= CONFIG =================

DOWNLOADS = Path.home() / "Downloads"

PASTA_BASE = Path(r"\\BRSBESRV960\Publico\REPORTS\ITAU\JOURNEY_LIBRAS\Pesquisa")
PASTA_BKP = PASTA_BASE / "BKP"

PADRAO_APP = "App - Pesquisas de Atendimento Atualizado*.xlsx"
PADRAO_AGENCIA = "AGENCIA - Pesquisas de Atendimento Atualizado*.xlsx"

NOME_FIXO_REDE = "PesquisaAtendimento.csv"
NOME_BASE_BKP = "PesquisaAtendimento"

LOG_DIR = Path("logs")
LOG_DIR.mkdir(exist_ok=True)

LOG_FILE = LOG_DIR / f"rpa_{datetime.now():%Y%m%d}.log"

# ==========================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout)
    ]
)

logger = logging.getLogger("RPA_PesquisaAtendimento")

# ==========================================


def localizar_arquivo_recente(padrao):
    arquivos = list(DOWNLOADS.glob(padrao))

    if not arquivos:
        logger.warning(f"Nenhum arquivo encontrado para padrão {padrao}")
        return None

    arquivo = max(arquivos, key=lambda f: f.stat().st_mtime)
    logger.info(f"Arquivo localizado: {arquivo.name}")
    return arquivo


def gerar_nome_bkp_incremental():
    contador = 1
    while True:
        nome = f"{NOME_BASE_BKP}_{contador}.csv"
        caminho = PASTA_BKP / nome
        if not caminho.exists():
            return caminho
        contador += 1


def mover_para_bkp():
    arquivo_atual = PASTA_BASE / NOME_FIXO_REDE

    if not arquivo_atual.exists():
        logger.info("Nenhum CSV anterior encontrado para backup.")
        return

    PASTA_BKP.mkdir(exist_ok=True)
    novo_nome = gerar_nome_bkp_incremental()
    shutil.move(arquivo_atual, novo_nome)
    logger.info(f"Backup criado: {novo_nome.name}")


def consolidar_excel_para_csv(app_path: Path, agencia_path: Path) -> Path:
    logger.info("Abrindo Excel para consolidação...")

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        # Abre os dois arquivos
        wb_app = excel.Workbooks.Open(str(app_path))
        wb_ag = excel.Workbooks.Open(str(agencia_path))

        ws_app = wb_app.Worksheets("Exportacao")
        ws_ag = wb_ag.Worksheets("Exportacao")

        # Remove filtros
        if ws_app.AutoFilterMode:
            ws_app.AutoFilterMode = False
        if ws_ag.AutoFilterMode:
            ws_ag.AutoFilterMode = False

        # Cria novo workbook consolidado
        wb_final = excel.Workbooks.Add()
        ws_final = wb_final.Worksheets(1)

        logger.info("Copiando dados do APP...")
        ws_app.UsedRange.Copy(ws_final.Range("A1"))

        # Descobre última linha usada após APP
        last_row = ws_final.Cells(ws_final.Rows.Count, 1).End(-4162).Row  # xlUp

        logger.info("Copiando dados da AGÊNCIA...")
        # Descobre limites da planilha AGÊNCIA
        last_row_ag = ws_ag.Cells(ws_ag.Rows.Count, 1).End(-4162).Row  # xlUp
        last_col_ag = ws_ag.Cells(1, ws_ag.Columns.Count).End(-4159).Column  # xlToLeft

        # Intervalo sem cabeçalho (começa na linha 2)
        range_ag = ws_ag.Range(
            ws_ag.Cells(2, 1),
            ws_ag.Cells(last_row_ag, last_col_ag)
        )

        range_ag.Copy(ws_final.Cells(last_row + 1, 1))
        caminho_csv = PASTA_BASE / NOME_FIXO_REDE

        if caminho_csv.exists():
            caminho_csv.unlink()

        wb_final.SaveAs(str(caminho_csv), FileFormat=62)  # CSV UTF-8
        wb_final.Close(False)

        logger.info("CSV consolidado gerado com sucesso.")

    finally:
        wb_app.Close(False)
        wb_ag.Close(False)
        excel.Quit()
        logger.info("Excel finalizado.")

    return caminho_csv


def main():
    inicio = time.time()
    logger.info("===== INÍCIO DO PROCESSO =====")

    try:
        arquivo_app = localizar_arquivo_recente(PADRAO_APP)
        arquivo_agencia = localizar_arquivo_recente(PADRAO_AGENCIA)

        if not arquivo_app or not arquivo_agencia:
            return

        mover_para_bkp()

        csv_final = consolidar_excel_para_csv(arquivo_app, arquivo_agencia)

        logger.info(f"Arquivo final disponível em: {csv_final}")

        tempo_total = round(time.time() - inicio, 2)
        logger.info(f"Processo finalizado com sucesso em {tempo_total}s")

    except Exception:
        logger.exception("Erro crítico durante execução do RPA.")

    finally:
        logger.info("===== FIM DO PROCESSO =====\n")


if __name__ == "__main__":
    main()