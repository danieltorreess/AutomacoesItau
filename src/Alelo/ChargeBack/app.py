import shutil
from pathlib import Path

from src.Alelo.ChargeBack.email_service import EmailServiceChargeBack
from src.Alelo.ChargeBack.downloader import ChargeBackDownloader
from src.Alelo.ChargeBack.processor_3580 import processar_3580
from src.Alelo.ChargeBack.processor_3547 import processar_3547


ASSUNTO_EMAIL = "RE: Atualização bases intercâmbio"

PASTA_TEMP = Path(r"C:\Temp\ChargeBack")

PASTA_REDE = Path(r"\\brsbesrv960\publico\REPORTS\ALELO\CHARGEBACK")


# ==========================================================
# 💾 SALVAR COM BKP INCREMENTAL
# ==========================================================

def salvar_com_bkp(caminho_origem: Path, destino_final: Path, subpasta_bkp: str):

    print(f"🔄 Preparando salvamento para: {destino_final.name}")

    if destino_final.exists():

        pasta_bkp = PASTA_REDE / "BCK" / subpasta_bkp
        pasta_bkp.mkdir(parents=True, exist_ok=True)

        contador = 1

        while True:
            nome_bkp = destino_final.stem + f"_{contador}.xlsx"
            caminho_bkp = pasta_bkp / nome_bkp

            if not caminho_bkp.exists():
                break

            contador += 1

        print(f"📦 Movendo arquivo antigo para BKP: {caminho_bkp.name}")
        shutil.move(destino_final, caminho_bkp)

    print("📤 Copiando novo arquivo para rede...")
    shutil.copy2(caminho_origem, destino_final)

    print("✅ Salvamento concluído.")



# ==========================================================
# 🚀 MAIN
# ==========================================================

def main():

    print("\n🚀 Iniciando RPA ChargeBack")
    print("===================================================")

    PASTA_TEMP.mkdir(parents=True, exist_ok=True)

    # Limpa arquivos antigos do TEMP
    for arquivo in PASTA_TEMP.glob("*"):
        try:
            arquivo.unlink()
        except Exception:
            pass

    email_service = EmailServiceChargeBack()
    downloader = ChargeBackDownloader(PASTA_TEMP)

    email = email_service.buscar_ultimo_email(ASSUNTO_EMAIL)

    if not email:
        print("⚠️ Nenhum e-mail encontrado.")
        return

    print(f"📧 Assunto: {email.Subject}")
    print(f"📅 Recebido em: {email.ReceivedTime}")
    print("---------------------------------------------------")

    anexos = downloader.salvar_todos_anexos(email)

    if not anexos:
        print("❌ Nenhum anexo .xlsx encontrado.")
        return

    print(f"📎 Total de anexos baixados: {len(anexos)}")
    print("===================================================")

    for arquivo in anexos:

        caminho = Path(arquivo)
        nome = caminho.name.lower()

        print(f"\n🔎 Analisando anexo: {caminho.name}")
        print("---------------------------------------------------")

        if "3580" in nome:

            print("➡ Identificado como 3580")

            arquivo_tratado = processar_3580(caminho)

            destino_final = PASTA_REDE / "Relatório_3580_Consolidado.xlsx"

            salvar_com_bkp(
                caminho_origem=arquivo_tratado,
                destino_final=destino_final,
                subpasta_bkp="3580"
            )

        elif "3547" in nome:

            print("➡ Identificado como 3547")

            arquivo_tratado = processar_3547(caminho)

            destino_raiz = PASTA_REDE / "Relatório_3547_Consolidado.xlsx"
            pasta_bkp_3547 = PASTA_REDE / "BCK" / "3547"
            pasta_bkp_3547.mkdir(parents=True, exist_ok=True)

            # Gerar nome incremental BKP
            contador = 1
            while True:
                nome_bkp = f"Relatório_3547_Consolidado_{contador}.xlsx"
                caminho_bkp = pasta_bkp_3547 / nome_bkp

                if not caminho_bkp.exists():
                    break

                contador += 1

            print(f"📦 Salvando BKP: {caminho_bkp.name}")
            shutil.copy2(arquivo_tratado, caminho_bkp)

            print("📤 Salvando arquivo na raiz para SSIS...")
            shutil.copy2(arquivo_tratado, destino_raiz)

            print("🗑 Limpando TEMP...")

        elif "dap" in nome:
            print("⚠️ DAP ainda não implementado.")

        else:
            print("⚠️ Anexo ignorado.")

    print("\n===================================================")
    print("🏁 Processo finalizado com sucesso.")


if __name__ == "__main__":
    main()
