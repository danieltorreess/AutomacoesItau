from pathlib import Path
from datetime import datetime

from src.Alelo.ServiceNow.email_service import EmailServiceServiceNow
from src.Alelo.ServiceNow.downloader import salvar_anexo


ASSUNTO = "ENC: Relatório de Chamados - Atento"

DESTINO_REDE = Path(r"\\brsbesrv960\publico\REPORTS\ALELO\SERVICENOW")


def main():

    print("\n🚀 Iniciando RPA ServiceNow")
    print("===================================================")

    email_service = EmailServiceServiceNow()

    print("🔎 Buscando e-mail no Outlook...")
    email = email_service.buscar_ultimo_email(ASSUNTO)

    if not email:
        print("⚠️ Nenhum e-mail encontrado.")
        return

    recebido_em = email.ReceivedTime
    remetente = email.SenderName
    assunto = email.Subject
    agora = datetime.now(recebido_em.tzinfo)

    diferenca = (agora - recebido_em).days

    print("📨 E-mail encontrado!")
    print(f"📧 Assunto: {assunto}")
    print(f"👤 Remetente: {remetente}")
    print(f"📅 Recebido em: {recebido_em}")
    print(f"⏳ Diferença em dias: {diferenca}")
    print(f"📎 Total de anexos: {email.Attachments.Count}")
    print("---------------------------------------------------")

    print("📥 Salvando anexo na rede...")

    arquivo_salvo = salvar_anexo(email, DESTINO_REDE)

    if not arquivo_salvo:
        print("❌ Nenhum anexo encontrado.")
        return

    print(f"✅ Arquivo salvo em: {arquivo_salvo}")
    print("🏁 Processo finalizado com sucesso.\n")


if __name__ == "__main__":
    main()
