from .email_service import RaeGovEmailService
from .processor import RaeGovProcessor


def main():
    print("\n🚀 Iniciando automação RAeGOV...\n")

    service = RaeGovEmailService()
    processor = RaeGovProcessor()

    emails = service.buscar_emails()

    if not emails:
        print("❌ Nenhum e-mail encontrado nos últimos 3 dias.")
        return

    email = emails[0]  # mais recente
    print(f"📨 Processando e-mail recebido em {email.ReceivedTime}")

    anexos = email.Attachments
    if anexos.Count == 0:
        print("⚠️ E-mail sem anexos.")
        return

    processor.processar_anexos(anexos)

    print("\n🎉 Processo RAeGOV concluído com sucesso!\n")


if __name__ == "__main__":
    main()
