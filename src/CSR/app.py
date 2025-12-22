from .email_service import EmailService
from .downloader import salvar_anexos

CAMINHO_REDE = r"\\BRSBESRV960\publico\REPORTS\ITAU\CSR IC\BASE DISCADOR"


def main():
    print("🚀 Iniciando automação de download das bases do discador...")

    # 1️⃣ Instancia o serviço de e-mail (TEM QUE VIR PRIMEIRO)
    email_service = EmailService()

    # 2️⃣ Busca os e-mails
    emails = email_service.buscar_emails_do_dia_anterior()

    # 3️⃣ Agora sim pega a data REAL usada
    data_ref = email_service.get_data_referencia()
    data_str = data_ref.strftime("%d.%m.%Y")

    print(f"🔎 Buscando e-mails do dia {data_str}...")

    if not emails:
        print("⚠️ Nenhum e-mail encontrado para a data de referência.")
        return

    print(f"📨 {len(emails)} e-mails encontrados. Extraindo anexos...")

    salvar_anexos(emails, CAMINHO_REDE)

    print("\n✅ Processo concluído com sucesso!")
    print(f"📁 Arquivos salvos em: {CAMINHO_REDE}")


if __name__ == "__main__":
    main()
