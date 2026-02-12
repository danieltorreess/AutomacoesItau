import os

from src.Alelo.Ouvidoria.email_service import EmailServiceAleloOuvidoria
from src.Alelo.Ouvidoria.downloader import AleloOuvidoriaDownloader
from src.Alelo.Ouvidoria.file_utils import mover_para_bkp
from src.Alelo.Ouvidoria.excel_processor import processar_arquivo_excel


# ===== CONFIG =====

PASTA_REDE = r"\\brsbesrv960\Publico\REPORTS\001 CARGAS\001 EXCEL\ALELO BKO\OUVIDORIA"
PASTA_BKP = r"\\brsbesrv960\Publico\REPORTS\001 CARGAS\001 EXCEL\ALELO BKO\OUVIDORIA\BCK\OuvidoriaBKP"

ASSUNTO_EMAIL = "BASE ATUALIZADA - OUVIDORIA"
PADRAO_ANEXO = "BASE ATUALIZADA OUVIDORIA*.xlsx"
NOME_FIXO = "BASE ATUALIZADA OUVIDORIA.xlsx"
NOME_BASE_BKP = "BASE ATUALIZADA OUVIDORIA"


def main():

    print("\n🚀 Iniciando RPA ALELO - OUVIDORIA")

    email_service = EmailServiceAleloOuvidoria()
    downloader = AleloOuvidoriaDownloader(PASTA_REDE)

    # --- Busca e-mail ---
    email = email_service.buscar_ultimo_email(ASSUNTO_EMAIL)

    if not email:
        print("⚠️ Nenhum e-mail encontrado.")
        return

    print(f"📨 E-mail encontrado: {email.Subject}")
    print(f"📅 Recebido em: {email.ReceivedTime}")

    # --- Backup do arquivo atual ---
    caminho_atual = os.path.join(PASTA_REDE, NOME_FIXO)

    mover_para_bkp(
        caminho_atual=caminho_atual,
        pasta_bkp=PASTA_BKP,
        nome_base=NOME_BASE_BKP
    )

    # --- Download do novo anexo ---
    novo_arquivo = downloader.salvar_anexo_padrao(
        email=email,
        padrao_nome=PADRAO_ANEXO,
        nome_final=NOME_FIXO
    )

    if not novo_arquivo:
        print("❌ Anexo não encontrado.")
        return

    print(f"📥 Arquivo salvo: {novo_arquivo}")

    # --- Processamento Excel ---
    processar_arquivo_excel(novo_arquivo)

    print("🎯 Processo finalizado com sucesso.")


if __name__ == "__main__":
    main()
