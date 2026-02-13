import win32com.client as win32
from datetime import datetime, timedelta


class EmailServiceBeedooUser:

    def __init__(self):
        self.outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

    def _get_folder(self):
        inbox = self.outlook.GetDefaultFolder(6)
        alelo = inbox.Folders["Alelo"]
        beedoo = alelo.Folders["Beedoo"]
        return beedoo

    def buscar_ultimo_email(self, assunto_base, dias=4):

        pasta = self._get_folder()
        mensagens = pasta.Items
        mensagens.Sort("[ReceivedTime]", True)

        hoje = datetime.now().date()
        limite = hoje - timedelta(days=dias)

        for msg in mensagens:
            if msg.Class != 43:
                continue

            if msg.ReceivedTime.date() < limite:
                break

            if assunto_base.lower() in (msg.Subject or "").lower():
                return msg

        return None
