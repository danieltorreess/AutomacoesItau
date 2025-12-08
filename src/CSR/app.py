from datetime import datetime, timedelta
from .email_service import EmailService
from .downloader import salvar_anexos

# Caminho onde os arquivos serão salvos
CAMINHO_REDE = r"\\BRSBESRV960\publico\REPORTS\ITAU\CSR IC\BASE DISCADOR"


def main():
    print("🚀 Iniciando automação de download das bases do discador...")

    # 1️⃣ Calcula a data do dia anterior
    referencia = datetime.today() - timedelta(days=1)
    data_str = referencia.strftime("%d.%m.%Y")
    print(f"🔎 Buscando e-mails do dia {data_str}...")

    # 2️⃣ Instancia o serviço de e-mail
    email_service = EmailService()

    # 3️⃣ Busca os e-mails do dia anterior
    emails = email_service.buscar_emails_do_dia_anterior()

    if not emails:
        print("⚠️ Nenhum e-mail encontrado para o dia anterior.")
        print("ℹ️ Verifique se as bases realmente foram enviadas ou se o nome do assunto mudou.")
        return

    print(f"📨 {len(emails)} e-mails encontrados. Extraindo anexos...")

    # 4️⃣ Salva anexos na rede
    salvar_anexos(emails, CAMINHO_REDE)

    print("\n✅ Processo concluído com sucesso!")
    print(f"📁 Arquivos salvos em: {CAMINHO_REDE}")


if __name__ == "__main__":
    main()
