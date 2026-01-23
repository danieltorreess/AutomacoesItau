import os

from src.OperacaoLibras.email_service import EmailServiceOperacaoLibras
from src.OperacaoLibras.downloader import OperacaoLibrasDownloader
from src.OperacaoLibras.file_utils import mover_para_bkp


# ===== CONFIGURAÇÕES =====

PASTA_BASE = r"\\BRSBESRV960\Publico\REPORTS\ITAU\JOURNEY_LIBRAS\Pesquisa"
PASTA_BKP = os.path.join(PASTA_BASE, "BKP")

ASSUNTO_EMAIL = (
    "Relatório de Atendimento / Pesquisas de Atendimento / Detratoras - Libras Itaú"
)

PADRAO_ANEXO = "Pesquisas de Atendimento atualizado*.xlsx"
NOME_FIXO_SAIDA = "PesquisaAtendimento.xlsx"
NOME_BASE_BKP = "PesquisaAtendimento"


def main():
    print("\n🚀 Iniciando RPA - Operação Libras | Pesquisa Atendimento")

    email_service = EmailServiceOperacaoLibras()
    downloader = OperacaoLibrasDownloader(PASTA_BASE)

    # --- Busca e-mail ---
    email = email_service.buscar_ultimo_email(ASSUNTO_EMAIL)

    if not email:
        print("⚠️ Nenhum e-mail encontrado.")
        return

    print(f"📨 Último e-mail recebido em {email.ReceivedTime}")

    # --- Backup do arquivo atual ---
    caminho_atual = os.path.join(PASTA_BASE, NOME_FIXO_SAIDA)

    mover_para_bkp(
        caminho_atual=caminho_atual,
        pasta_bkp=PASTA_BKP,
        nome_base=NOME_BASE_BKP
    )

    # --- Download do novo anexo ---
    novo_arquivo = downloader.salvar_anexo_padrao(
        email=email,
        padrao_nome=PADRAO_ANEXO,
        nome_final=NOME_FIXO_SAIDA
    )

    if not novo_arquivo:
        print("❌ Anexo correto não encontrado no e-mail.")
        return

    print(f"✅ Novo arquivo salvo: {novo_arquivo}")
    print("🎯 Arquivo pronto para consumo pelo SSIS")


if __name__ == "__main__":
    main()
