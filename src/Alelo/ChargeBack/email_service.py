import win32com.client as win32
from datetime import datetime, timedelta


class EmailServiceChargeBack:

    def __init__(self):
        print("🔌 Conectando ao Outlook...")
        self.outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

    def _get_folder(self):
        inbox = self.outlook.GetDefaultFolder(6)
        alelo = inbox.Folders["Alelo"]
        chargeback = alelo.Folders["ChargeBack"]
        return chargeback

    def buscar_ultimo_email(self, assunto_alvo, dias=4):

        pasta = self._get_folder()
        mensagens = pasta.Items
        mensagens.Sort("[ReceivedTime]", True)

        hoje = datetime.now().date()
        limite = hoje - timedelta(days=dias)

        print(f"📅 Buscando e-mails dos últimos {dias} dias...")

        for msg in mensagens:

            if msg.Class != 43:
                continue

            data_email = msg.ReceivedTime.date()

            if data_email < limite:
                break

            if assunto_alvo.lower() in (msg.Subject or "").lower():
                print("📨 E-mail encontrado!")
                return msg

        return None
