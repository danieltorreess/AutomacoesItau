# src/BKO/app.py
from .email_service import EmailServiceBKO
from .processor import BKOProcessor
import os


def main():
    print("\n🚀 Iniciando automação BKO...\n")

    email_service = EmailServiceBKO()
    processor = BKOProcessor()

    # -------------------------------------------------------
    # 📌 1. Pós Venda
    # -------------------------------------------------------
    print("🔎 Buscando e-mail Pós Venda do dia...")

    msg_pos = email_service.buscar_email_pos_venda()

    if msg_pos:
        print(f"📨 Encontrado: {msg_pos.Subject}")

        for att in msg_pos.Attachments:
            name = att.FileName.lower()
            if name.endswith(".csv") and "pós venda" in name.replace("pos", "pós"):
                print(f"📁 Processando arquivo Pós Venda: {att.FileName}")
                path = processor.processar_pos_venda(att, att.FileName)
                print(f"✅ Pós Venda salvo em: {path}\n")
                break
    else:
        print("⚠️ Nenhum e-mail Pós Venda encontrado para HOJE.\n")

    # -------------------------------------------------------
    # 📌 2. Segundo nível atendimento PJ
    # -------------------------------------------------------
    print("🔎 Buscando e-mail 2º nível do dia...")

    msg_seg = email_service.buscar_email_segundo_nivel()

    if msg_seg:
        print(f"📨 Encontrado: {msg_seg.Subject}")

        for att in msg_seg.Attachments:
            name = att.FileName.lower()
            if name.endswith(".csv") and "preenchimento" in name:
                print(f"📁 Processando base 2º nível: {att.FileName}")
                path = processor.processar_segundo_nivel(att, att.FileName)
                print(f"✅ 2º nível salvo em: {path}\n")
                break
    else:
        print("⚠️ Nenhum e-mail 2º nível encontrado para HOJE.\n")

    print("🎉 Automação BKO concluída!\n")


if __name__ == "__main__":
    main()
