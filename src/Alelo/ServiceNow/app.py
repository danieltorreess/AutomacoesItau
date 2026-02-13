from pathlib import Path

from src.Alelo.ServiceNow.email_service import EmailServiceServiceNow
from src.Alelo.ServiceNow.downloader import salvar_anexo


ASSUNTO = "ENC: Relatório de Chamados - Atento"

DESTINO_REDE = Path(r"\\brsbesrv960\publico\REPORTS\ALELO\SERVICENOW")


def main():

    print("🚀 Iniciando RPA ServiceNow")

    email_service = EmailServiceServiceNow()

    email = email_service.buscar_ultimo_email(ASSUNTO)

    if not email:
        print("⚠️ Nenhum e-mail encontrado.")
        return

    print(f"📨 E-mail encontrado: {email.Subject}")

    arquivo_salvo = salvar_anexo(email, DESTINO_REDE)

    if not arquivo_salvo:
        print("❌ Nenhum anexo encontrado.")
        return

    print(f"✅ Arquivo salvo em: {arquivo_salvo}")
    print("🏁 Processo finalizado com sucesso.")


if __name__ == "__main__":
    main()
