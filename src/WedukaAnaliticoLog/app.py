import os

from src.WedukaAnaliticoLog.email_service import EmailServiceWeduka
from src.WedukaAnaliticoLog.downloader import WedukaDownloader
from src.WedukaAnaliticoLog.file_utils import mover_para_bkp


# ===== CONFIGURAÇÕES =====

PASTA_ANALITICO = r"\\BRSBESRV960\Publico\REPORTS\ITAU\WEDUKA_TREINAMENTO\AnaliticoDiario"
PASTA_BKP = os.path.join(PASTA_ANALITICO, "BKP")

# 🔹 Assuntos aceitos (antigo + novos)
ASSUNTOS_EMAIL = [
    "LOG DE ACESSO DIARIO WEDUKA",
    "Relatórios de Log Diários - Weduka",
    "Relatório de Log Diário - Weduka",
]

PADRAO_ANEXO = "Analitico_Log_de_acesso_diario*.xlsx"
NOME_FIXO_SAIDA = "Analitico_Log_de_acesso_diario.xlsx"
NOME_BASE_BKP = "Analitico_Log_de_acesso_diario"


def main():
    print("\n🚀 Iniciando RPA WEDUKA - LOG DE ACESSO DIÁRIO")

    email_service = EmailServiceWeduka()
    downloader = WedukaDownloader(PASTA_ANALITICO)

    # --- Busca e-mail ---
    email = email_service.buscar_ultimo_email(ASSUNTOS_EMAIL)

    if not email:
        print("⚠️ Nenhum e-mail encontrado para os assuntos configurados.")
        return

    print(f"📨 Último e-mail recebido em {email.ReceivedTime}")
    print(f"📌 Assunto: {email.Subject}")

    # --- Backup do arquivo atual ---
    caminho_atual = os.path.join(PASTA_ANALITICO, NOME_FIXO_SAIDA)

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
        print("❌ Anexo não encontrado no e-mail.")
        return

    print(f"✅ Novo arquivo salvo: {novo_arquivo}")
    print("🎯 Arquivo pronto para consumo pelo SSIS")


if __name__ == "__main__":
    main()
