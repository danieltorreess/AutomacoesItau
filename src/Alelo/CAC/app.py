from pathlib import Path
import time

from src.Alelo.CAC.email_service import EmailServiceCAC
from src.Alelo.CAC.processor import processar_excel
from src.Alelo.CAC.file_utils import mover_para_bkp


ASSUNTO = "ENC: Relatar resultados (Status das Demandas | Operações CAC)"

DESTINO_REDE = Path(r"\\brsbesrv960\publico\REPORTS\ALELO\CAC - EC VAREJO")

PASTA_BKP = DESTINO_REDE / "BKP"

NOME_FINAL = "Status das Demandas   Operações CAC.xlsx"


def salvar_anexo(email, pasta_temp: Path):

    pasta_temp.mkdir(parents=True, exist_ok=True)

    for att in email.Attachments:

        nome = att.FileName.lower()

        # 🔥 Só aceita arquivos Excel
        if nome.endswith(".xlsx") or nome.endswith(".xls"):

            caminho = pasta_temp / att.FileName
            att.SaveAsFile(str(caminho))
            return caminho

    return None


def main():

    print("🚀 Iniciando RPA CAC")

    email_service = EmailServiceCAC()
    email = email_service.buscar_ultimo_email(ASSUNTO)

    if not email:
        print("⚠️ Nenhum e-mail encontrado.")
        return

    print(f"📨 E-mail encontrado: {email.Subject}")

    # ==============================
    # 📥 Salvar anexo temporariamente
    # ==============================

    pasta_temp = Path.home() / "Downloads"
    arquivo_origem = salvar_anexo(email, pasta_temp)

    if not arquivo_origem:
        print("❌ Nenhum anexo encontrado.")
        return

    print(f"📥 Anexo salvo: {arquivo_origem}")

    caminho_destino = DESTINO_REDE / NOME_FINAL

    # ==============================
    # 📦 Backup
    # ==============================

    mover_para_bkp(caminho_destino, PASTA_BKP)

    # ==============================
    # 🔄 Processar e padronizar
    # ==============================

    processar_excel(arquivo_origem, caminho_destino)

    # ==============================
    # 🗑 Remover arquivo temporário
    # ==============================

    for tentativa in range(3):
        try:
            time.sleep(1)
            arquivo_origem.unlink()
            break
        except PermissionError:
            time.sleep(2)

    print("✅ Processo CAC finalizado com sucesso.")


if __name__ == "__main__":
    main()
