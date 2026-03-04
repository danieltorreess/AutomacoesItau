from .email_service import EmailService
from .attachment_downloader import AttachmentDownloader
from .backup_service import mover_para_bkp
from .logger import log


def main():

    log("🚀 Iniciando RPA GENESYS")

    email_service = EmailService()

    downloader = AttachmentDownloader()

    emails = email_service.buscar_emails()

    log(f"📨 {len(emails)} emails encontrados")

    for email in emails:

        log(f"\n📩 Email: {email.Subject}")

        log(f"🕒 Recebido: {email.ReceivedTime}")

        arquivos = downloader.salvar_anexos(email)

        for tipo, caminho in arquivos:

            mover_para_bkp(tipo, caminho)

    log("\n🏁 Processo finalizado")


if __name__ == "__main__":
    main()