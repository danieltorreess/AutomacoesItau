from pathlib import Path

from src.Alelo.BasePBI.email_service import EmailServiceAleloBasePBI
from src.Alelo.BasePBI.link_extractor import extrair_link_download
from src.Alelo.BasePBI.downloader import baixar_arquivo
from src.Alelo.BasePBI.file_utils import ajustar_nome_arquivo, copiar_para_destinos


ASSUNTO_BASE = "ENC: [Alelo] Base PBI Atento - Dados_Satisfacao | Acessos_Menu | Dados_Atendimento"

PASTA_TEMP = r"C:\Users\ab1541240\Downloads"

DESTINOS = [
    r"\\brsbesrv960\publico\REPORTS\ALELO\BLIP",
    r"C:\Users\ab1541240\OneDrive - ATENTO Global\Documents\ALELO\RelatoriosAlelo\AtualizacoesAlelo\HistoricoFerias"
]


def main():

    print("\n🚀 Iniciando RPA ALELO - BasePBI")

    email_service = EmailServiceAleloBasePBI()

    email = email_service.buscar_ultimo_email(ASSUNTO_BASE)

    if not email:
        print("⚠️ Nenhum e-mail encontrado.")
        return

    print(f"📨 E-mail encontrado: {email.Subject}")

    # corpo = email.Body

    # link = extrair_link_download(corpo)

    html = email.HTMLBody
    link = extrair_link_download(html)


    if not link:
        print("❌ Link de download não encontrado.")
        return

    print(f"🔗 Link encontrado: {link}")

    # --- Download ---
    arquivo_baixado = baixar_arquivo(link, PASTA_TEMP)

    print(f"⬇ Arquivo baixado: {arquivo_baixado}")

    # --- Ajusta nome ---
    arquivo_renomeado = ajustar_nome_arquivo(arquivo_baixado)

    print(f"✏ Nome ajustado: {arquivo_renomeado.name}")

    # --- Copia para destinos ---
    copiar_para_destinos(arquivo_renomeado, DESTINOS)

    print("✅ Processo finalizado com sucesso.")


if __name__ == "__main__":
    main()
