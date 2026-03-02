from pathlib import Path
from datetime import datetime

from src.Alelo.BKOPrevencao.email_service import EmailServiceBKOPrevencao
from src.Alelo.BKOPrevencao.downloader import salvar_anexo
from src.Alelo.BKOPrevencao.excel_processor import processar_excel
from src.Alelo.BKOPrevencao.file_manager import backup_arquivo

import time
import gc


ASSUNTO = "Base Mesas - MIS31041 - Intraday - BackOffice - Prevenção"

DESTINO_REDE = Path(
    r"\\brsbesrv960\publico\REPORTS\001 CARGAS\001 EXCEL\ALELO\PREVENCAO"
)

PASTA_BKP = DESTINO_REDE / "BKP"
ARQUIVO_FINAL = DESTINO_REDE / "BASE.xlsx"

BASE_DIR = Path(__file__).resolve().parents[3]
PASTA_TEMP = BASE_DIR / "temp"

def main():

    print("\n🚀 Iniciando RPA BKO Prevenção")
    print("=" * 60)

    email_service = EmailServiceBKOPrevencao()

    print("🔎 Buscando e-mail no Outlook...")
    email = email_service.buscar_ultimo_email(ASSUNTO)

    if not email:
        print("❌ Nenhum e-mail encontrado.")
        return

    recebido_em = email.ReceivedTime
    print("📨 E-mail encontrado!")
    print(f"📧 Assunto: {email.Subject}")
    print(f"👤 Remetente: {email.SenderName}")
    print(f"📅 Data/Hora recebimento: {recebido_em}")
    print(f"📎 Anexos: {email.Attachments.Count}")
    print("-" * 60)

    print("📥 Salvando anexo temporariamente...")
    arquivo_anexo = salvar_anexo(email, PASTA_TEMP)

    if not arquivo_anexo:
        print("❌ Nenhum anexo .xlsx encontrado.")
        return

    print(f"📁 Anexo salvo em: {arquivo_anexo}")

    print("📦 Executando backup do arquivo atual...")
    backup_arquivo(ARQUIVO_FINAL, PASTA_BKP)

    print("⚙️ Processando Excel...")
    processar_excel(arquivo_anexo, ARQUIVO_FINAL)

    print("🧹 Limpando arquivos temporários...")

    gc.collect()

    for tentativa in range(5):
        try:
            arquivo_anexo.unlink()
            print("✅ Arquivo temporário removido.")
            break
        except PermissionError:
            print(f"⚠ Arquivo ainda em uso... tentativa {tentativa+1}/5")
            time.sleep(1)
    else:
        print("❌ Não foi possível remover o arquivo temporário.")
        print("🏁 Processo finalizado com sucesso.\n")


if __name__ == "__main__":
    main()