from pathlib import Path

from src.Alelo.BeedooLocal.email_service import EmailServiceBeedooUser
from src.Alelo.BeedooLocal.downloader_email import salvar_anexo_user
from src.Alelo.BeedooLocal.processor import processar_excel_para_csv
from src.Alelo.BeedooLocal.file_utils import mover_para_bkp


# ==========================
# CONFIG
# ==========================

ASSUNTO = "ENC: Relatório de acessos Beedoo"

DOWNLOAD_DIR = Path.home() / "Downloads"

DESTINO_REDE = Path(r"\\brsbesrv960\publico\REPORTS\ALELO\BEEDOO")

BKP_USER = DESTINO_REDE / "BKP" / "user"

DESTINO_REDE.mkdir(parents=True, exist_ok=True)


def main():

    print("🚀 Iniciando RPA BEEDOO - USER (Email)")

    email_service = EmailServiceBeedooUser()

    email = email_service.buscar_ultimo_email(ASSUNTO)

    if not email:
        print("⚠️ Nenhum e-mail encontrado.")
        return

    print(f"📨 E-mail encontrado: {email.Subject}")

    # ==============================
    # 📥 Baixar anexo
    # ==============================

    arquivo_xlsx = salvar_anexo_user(email, DOWNLOAD_DIR)

    if not arquivo_xlsx:
        print("❌ Anexo user não encontrado.")
        return

    print(f"📥 Anexo salvo: {arquivo_xlsx}")

    # ==============================
    # 📦 Backup CSV existente
    # ==============================

    destino_csv = DESTINO_REDE / "user.csv"

    mover_para_bkp(destino_csv, BKP_USER)

    # ==============================
    # 🔄 Converter XLSX → CSV
    # ==============================

    processar_excel_para_csv(
        arquivo_xlsx,
        "user",
        DESTINO_REDE
    )

    # ==============================
    # 🗑 Remove XLSX temporário
    # ==============================

    arquivo_xlsx.unlink()

    print("✅ Processo USER finalizado com sucesso.")


if __name__ == "__main__":
    main()
