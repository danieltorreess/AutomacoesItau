import win32com.client as win32
from datetime import datetime
from .config import EMAIL_SUBJECT, DATA_INICIAL, DATA_FINAL


class EmailService:

    def __init__(self):
        self.outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

    def _get_folder(self):
        inbox = self.outlook.GetDefaultFolder(6)
        genesys = inbox.Folders["GENESYS"]
        return genesys

    def buscar_emails(self):

        pasta = self._get_folder()
        emails = []

        for msg in pasta.Items:

            if msg.Class != 43:
                continue

            subject = msg.Subject or ""

            if EMAIL_SUBJECT.lower() not in subject.lower():
                continue

            data_email = msg.ReceivedTime.date()

            if data_email < DATA_INICIAL:
                continue

            if data_email > DATA_FINAL:
                continue

            emails.append(msg)

        emails.sort(key=lambda x: x.ReceivedTime)

        return emails