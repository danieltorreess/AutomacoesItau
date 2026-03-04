import win32com.client as win32
from .config import EMAIL_SUBJECTS
from .date_utils import datas_para_busca


class EmailService:

    def __init__(self):

        self.outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

    def get_folder(self):

        inbox = self.outlook.GetDefaultFolder(6)

        return inbox.Folders["GENESYS"]

    def buscar_emails(self):

        pasta = self.get_folder()

        datas_validas = datas_para_busca()

        emails_validos = []

        for msg in pasta.Items:

            if msg.Class != 43:
                continue

            assunto = msg.Subject or ""

            if not any(a.lower() in assunto.lower() for a in EMAIL_SUBJECTS):
                continue

            data_email = msg.ReceivedTime.date()

            if data_email not in datas_validas:
                continue

            emails_validos.append(msg)

        emails_validos.sort(key=lambda x: x.ReceivedTime)

        return emails_validos