from pathlib import Path
from datetime import datetime, timedelta

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


def obter_datas_para_processar():
    hoje = datetime.now().date()
    dia_semana = hoje.weekday()  # 0 = segunda

    if dia_semana == 0:
        # Segunda-feira → sábado, domingo e segunda
        return [
            hoje - timedelta(days=2),
            hoje - timedelta(days=1),
            hoje
        ]
    else:
        # Terça a sexta → apenas hoje
        return [hoje]


def main():

    print("\n🚀 Iniciando RPA ALELO - BasePBI")

    email_service = EmailServiceAleloBasePBI()

    datas_para_processar = obter_datas_para_processar()

    print(f"📅 Datas alvo: {datas_para_processar}")

    emails = email_service.buscar_emails_por_datas(
        ASSUNTO_BASE,
        datas_para_processar
    )

    if not emails:
        print("⚠️ Nenhum e-mail encontrado para as datas alvo.")
        return

    for email in emails:

        print(f"\n📨 Processando e-mail: {email.Subject}")
        print(f"📅 Data: {email.ReceivedTime.date()}")

        html = email.HTMLBody
        link = extrair_link_download(html)

        if not link:
            print("❌ Link de download não encontrado.")
            continue

        print(f"🔗 Link encontrado: {link}")

        # Download
        arquivo_baixado = baixar_arquivo(link, PASTA_TEMP)
        print(f"⬇ Arquivo baixado: {arquivo_baixado}")

        # Ajuste de nome
        arquivo_renomeado = ajustar_nome_arquivo(arquivo_baixado)
        print(f"✏ Nome ajustado: {arquivo_renomeado.name}")

        # Copiar para destinos
        copiar_para_destinos(arquivo_renomeado, DESTINOS)

    print("\n✅ Processo finalizado com sucesso.")


if __name__ == "__main__":
    main()
