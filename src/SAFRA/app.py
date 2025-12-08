import os
from .email_service import EmailServiceSafra
from .downloader import SafraDownloader
from .excel_utils import desocultar_abas

OUTPUT = r"\\BRSBESRV960\Publico\REPORTS\SAFRA\SHRINKAGE"

ASSUNTO_GERENCIAL = "GERENCIAL_LOG E PA"
ASSUNTO_MIS = "MIS31047 - MICRO-GESTÃO_ATENDIMENTO_CONSIGNADO E AGZERO"


def processar_gerencial(email_service, downloader):
    print("\n🔎 Buscando último e-mail GERENCIAL_LOG E PA...")

    email = email_service.buscar_ultimo_email(ASSUNTO_GERENCIAL)
    if not email:
        print("⚠️ Nenhum e-mail GERENCIAL encontrado.")
        return

    print(f"📨 Último GERENCIAL recebido em {email.ReceivedTime}")

    caminho = downloader.salvar_anexo_excel(email)
    if not caminho:
        print("❌ Nenhum anexo Excel encontrado no GERENCIAL.")
        return

    print(f"📁 Arquivo salvo: {caminho}")

    print("🛠 Desocultando todas as abas...")
    desocultar_abas(caminho)

    print("✅ GERENCIAL_LOG E PA processado com sucesso!")


def processar_mis(email_service, downloader):
    print("\n🔎 Buscando último e-mail MIS31047...")

    email = email_service.buscar_ultimo_email(ASSUNTO_MIS)
    if not email:
        print("⚠️ Nenhum e-mail MIS encontrado.")
        return

    print(f"📨 Último MIS recebido em {email.ReceivedTime}")

    caminho = downloader.salvar_anexo_excel(email)
    if not caminho:
        print("❌ Nenhum anexo Excel encontrado no MIS.")
        return

    print(f"📁 Arquivo salvo: {caminho}")
    print("✅ MIS31047 processado com sucesso (nenhuma transformação necessária).")


def main():
    print("\n🚀 Iniciando automação SAFRA...")

    email_service = EmailServiceSafra()
    downloader = SafraDownloader(OUTPUT)

    # ----- FLUXO 1 -----
    processar_gerencial(email_service, downloader)

    # ----- FLUXO 2 -----
    processar_mis(email_service, downloader)

    print("\n🎉 Automação SAFRA concluída com sucesso!")


if __name__ == "__main__":
    main()
