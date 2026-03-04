import os

from .email_service import EmailService
from .attachment_downloader import AttachmentDownloader
from .consolidator import Consolidator
from .logger import log

from .config import (
    PASTA_AGENDAMENTO,
    PASTA_AUTORIZACAO,
    OUTPUT_FINAL,
    CONSOLIDADO_AGENDAMENTO,
    CONSOLIDADO_AUTORIZACAO
)


def main():

    log("🚀 Iniciando RPA Genesys")

    email_service = EmailService()
    downloader = AttachmentDownloader()
    consolidator = Consolidator()

    emails = email_service.buscar_emails()

    log(f"📨 {len(emails)} e-mails encontrados")

    for email in emails:

        log(f"📩 Processando email {email.ReceivedTime}")

        arquivos = downloader.salvar_anexos(email)

        for arq in arquivos:
            log(f"💾 Salvo: {arq}")

    log("📊 Iniciando consolidação AGENDAMENTO")

    consolidator.consolidar(
        PASTA_AGENDAMENTO,
        os.path.join(OUTPUT_FINAL, CONSOLIDADO_AGENDAMENTO)
    )

    log("📊 Iniciando consolidação AUTORIZAÇÃO")

    consolidator.consolidar(
        PASTA_AUTORIZACAO,
        os.path.join(OUTPUT_FINAL, CONSOLIDADO_AUTORIZACAO)
    )

    log("🏁 Processo finalizado")


if __name__ == "__main__":
    main()