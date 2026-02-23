import shutil
from pathlib import Path

from src.Alelo.ChargeBack.email_service import EmailServiceChargeBack
from src.Alelo.ChargeBack.downloader import ChargeBackDownloader
from src.Alelo.ChargeBack.processor_3580 import processar_3580
from src.Alelo.ChargeBack.processor_3547 import processar_3547
from src.Alelo.ChargeBack.processor_dap import processar_dap


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

    # ==========================================================
    # 🔎 FASE 1 - IDENTIFICAÇÃO INICIAL
    # ==========================================================

    print("\n🧠 Iniciando identificação inteligente dos anexos...")
    print("===================================================")

    arquivos_processados = set()

    print(f"\n📦 Total de anexos recebidos: {len(anexos)}")

    for arquivo in anexos:

        caminho = Path(arquivo)
        nome = caminho.name.lower()

        print(f"\n🔎 Analisando anexo: {caminho.name}")
        print("---------------------------------------------------")

        # ------------------------------------------------------
        # 3580
        # ------------------------------------------------------
        if "3580" in nome:

            print("🟢 Regra: nome contém '3580'")
            print("➡ Classificado como BASE 3580")

            arquivo_tratado = processar_3580(caminho)

            destino_final = PASTA_REDE / "Relatório_3580_Consolidado.xlsx"

            salvar_com_bkp(
                caminho_origem=arquivo_tratado,
                destino_final=destino_final,
                subpasta_bkp="3580"
            )

            arquivos_processados.add(arquivo)
            continue

        # ------------------------------------------------------
        # DAP
        # ------------------------------------------------------
        if "dap" in nome:

            print("🟢 Regra: nome contém 'dap'")
            print("➡ Classificado como BASE DAP")

            arquivo_tratado = processar_dap(caminho)

            pasta_dap = PASTA_REDE / "DAP"
            pasta_bkp = pasta_dap / "BCK"

            pasta_dap.mkdir(parents=True, exist_ok=True)
            pasta_bkp.mkdir(parents=True, exist_ok=True)

            destino_final = pasta_dap / caminho.name

            contador = 1
            while True:
                nome_bkp = caminho.stem + f"_{contador}.xlsx"
                caminho_bkp = pasta_bkp / nome_bkp

                if not caminho_bkp.exists():
                    break

                contador += 1

            print(f"📦 BKP gerado: {caminho_bkp.name}")
            shutil.copy2(arquivo_tratado, caminho_bkp)

            print("📤 Arquivo DAP salvo na raiz.")
            shutil.copy2(arquivo_tratado, destino_final)

            arquivos_processados.add(arquivo)
            continue

        # ------------------------------------------------------
        # Ainda não classificado
        # ------------------------------------------------------
        print("🟡 Nenhuma regra direta aplicada (não é 3580 nem DAP).")
        print("🔎 Arquivo ficará aguardando decisão como possível 3547.")

    # ==========================================================
    # 🔥 FASE 2 - DEFINIÇÃO AUTOMÁTICA DO 3547
    # ==========================================================

    print("\n🧠 Iniciando verificação de arquivos restantes...")
    print("===================================================")

    arquivos_restantes = [a for a in anexos if a not in arquivos_processados]

    print(f"📌 Arquivos já processados: {len(arquivos_processados)}")
    print(f"📌 Arquivos restantes: {len(arquivos_restantes)}")

    if len(arquivos_restantes) == 1:

        caminho = Path(arquivos_restantes[0])

        print("\n🟢 Regra de negócio aplicada:")
        print("➡ Arquivo restante será tratado como BASE 3547")
        print(f"📄 Arquivo identificado: {caminho.name}")
        print("---------------------------------------------------")

        arquivo_tratado = processar_3547(caminho)

        destino_raiz = PASTA_REDE / "Relatório_3547_Consolidado.xlsx"
        pasta_bkp_3547 = PASTA_REDE / "BCK" / "3547"
        pasta_bkp_3547.mkdir(parents=True, exist_ok=True)

        contador = 1
        while True:
            nome_bkp = f"Relatório_3547_Consolidado_{contador}.xlsx"
            caminho_bkp = pasta_bkp_3547 / nome_bkp

            if not caminho_bkp.exists():
                break

            contador += 1

        print(f"📦 BKP gerado: {caminho_bkp.name}")
        shutil.copy2(arquivo_tratado, caminho_bkp)

        print("📤 Arquivo 3547 salvo na raiz para SSIS.")
        shutil.copy2(arquivo_tratado, destino_raiz)

    elif len(arquivos_restantes) == 0:

        print("\nℹ️ Nenhum arquivo restante.")
        print("➡ Apenas DAP e 3580 foram recebidos.")
        print("✔ Processo segue normalmente.")

    else:

        print("\n🔴 ALERTA DE INCONSISTÊNCIA")
        print("Mais de um arquivo restante encontrado.")
        print("Isso foge da regra esperada (máximo 3 anexos).")
        print("\nArquivos pendentes:")

        for a in arquivos_restantes:
            print(f"- {Path(a).name}")

        print("\n⚠️ Verifique manualmente.")

    print("\n===================================================")
    print("🏁 Processo finalizado com sucesso.")

if __name__ == "__main__":
    main()
