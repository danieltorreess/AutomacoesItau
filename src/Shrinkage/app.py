from datetime import datetime
import os

from .email_service import EmailServiceShrinkage
from .downloader import ShrinkageDownloader
from .atendimento_processor import AtendimentoProcessor


# ===== CONFIGURAÇÕES =====
OUTPUT_PATH = r"\\BRSBESRV960\Publico\REPORTS\ITAU\SHRINKAGE"
ESPELHO_TESTE = os.path.join(OUTPUT_PATH, "HUB_ATENDIMENTO_ATT_NEW.xlsx")


def main():
    print("\n🚀 Iniciando automação Shrinkage...\n")

    # ======================================================
    # 1️⃣ DEFINIR DATA DE REFERÊNCIA (SEMANA ATUAL)
    # ======================================================
    referencia = datetime.today()
    data_str = referencia.strftime("%d/%m/%Y")

    print(f"🔎 Buscando e-mails do dia {data_str}...")

    # ======================================================
    # 2️⃣ BUSCAR E-MAILS NA PASTA SHRINKAGE
    # ======================================================
    email_service = EmailServiceShrinkage()
    emails = email_service.buscar_emails_ultimos_dias(dias=3)

    if not emails:
        print("⚠️ Nenhum e-mail encontrado hoje na pasta Shrinkage.")
        return

    print(f"📨 Total de e-mails encontrados: {len(emails)}\n")

    # ======================================================
    # 3️⃣ EXTRAIR ANEXOS ATT-YYYYMM
    # ======================================================
    print("📥 Extraindo anexos ATT-* dos e-mails...")

    downloader = ShrinkageDownloader(OUTPUT_PATH)
    arquivos_att = downloader.processar_emails(emails)

    if not arquivos_att:
        print("⚠️ Nenhum arquivo ATT-* foi encontrado nos e-mails.")
        return

    print("📁 Arquivos ATT extraídos:")
    for nome, caminho in arquivos_att.items():
        print(f"   → {nome}")

    # ======================================================
    # 4️⃣ IDENTIFICAR O ARQUIVO DE HUB ATENDIMENTO
    # ======================================================
    hub_atendimento = None

    for nome, caminho in arquivos_att.items():
        if "atendimento" in nome.lower():
            hub_atendimento = caminho
            break

    if not hub_atendimento:
        print("\n❌ Não foi encontrado o arquivo ATT-* HUB Atendimento.")
        return

    print(f"\n🟦 Arquivo de Atendimento identificado:\n   → {hub_atendimento}")

    # ======================================================
    # 5️⃣ PROCESSAR O ARQUIVO (VOZ + DIGITAL)
    # ======================================================
    print("\n🛠 Iniciando processamento das tabelas dinâmicas VOZ e DIGITAL...")

    processor = AtendimentoProcessor()
    sucesso = processor.processar(hub_atendimento, ESPELHO_TESTE)

    if not sucesso:
        print("\n❌ Ocorreu um erro no processamento do HUB Atendimento.")
        return

    # ======================================================
    # 6️⃣ FINALIZAÇÃO
    # ======================================================
    print("\n✅ Processo Shrinkage finalizado com sucesso!")
    print(f"📁 Arquivo New gerado:\n   → {ESPELHO_TESTE}\n")


if __name__ == "__main__":
    main()
