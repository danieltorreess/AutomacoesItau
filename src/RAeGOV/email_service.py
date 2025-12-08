import win32com.client as win32
from datetime import datetime, timedelta


class RaeGovEmailService:

    def __init__(self):
        self.outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

    def _folder(self):
        inbox = self.outlook.GetDefaultFolder(6)  # Inbox
        itau = inbox.Folders["Itau"]
        return itau.Folders["RAeGOV"]

    def buscar_emails(self):
        folder = self._folder()

        hoje = datetime.today().date()
        limite = hoje - timedelta(days=3)

        emails_validos = []

        for msg in folder.Items:
            if msg.Class != 43:
                continue

            assunto = msg.Subject.lower()
            if "consignado - canais críticos" not in assunto:
                continue

            data_email = msg.ReceivedTime.date()
            if data_email < limite:
                continue

            emails_validos.append(msg)

        # ordenar por recebido mais recente
        emails_validos.sort(key=lambda m: m.ReceivedTime, reverse=True)

        return emails_validos
