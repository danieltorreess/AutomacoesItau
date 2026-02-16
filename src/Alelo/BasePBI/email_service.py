import win32com.client as win32
from datetime import datetime


class EmailServiceAleloBasePBI:

    def __init__(self):
        self.outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

    def _get_folder(self):
        inbox = self.outlook.GetDefaultFolder(6)
        alelo = inbox.Folders["Alelo"]
        blip = alelo.Folders["BLIP"]
        return blip

    def buscar_emails_por_datas(self, assunto_base, datas_alvo):
        """
        Retorna lista de emails cujo assunto contenha assunto_base
        e cuja data esteja dentro das datas_alvo.
        """

        pasta = self._get_folder()
        mensagens = pasta.Items
        mensagens.Sort("[ReceivedTime]", True)

        emails_encontrados = []
        data_minima = min(datas_alvo)

        for msg in mensagens:

            # Apenas MailItem
            if msg.Class != 43:
                continue

            data_email = msg.ReceivedTime.date()

            # Como está ordenado DESC, se passou da menor data pode parar
            if data_email < data_minima:
                break

            if data_email in datas_alvo:
                if assunto_base.lower() in (msg.Subject or "").lower():
                    emails_encontrados.append(msg)

        return emails_encontrados
