import win32com.client as win32
from datetime import datetime, timedelta


class EmailServiceShrinkage:
    """
    Serviço responsável por buscar os e-mails encaminhados do Shrinkage.
    """

    def __init__(self):
        self.outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

    def _get_target_folder(self):
        """
        Navega para:
        Caixa de Entrada > Itau > Shrinkage
        """
        inbox = self.outlook.GetDefaultFolder(6)  # Inbox
        itau = inbox.Folders["Itau"]
        shrinkage = itau.Folders["Shrinkage"]
        return shrinkage

    # ---------------------------------------------------------
    # MÉTODO INTERNO COMUM PARA BUSCAR POR UMA DATA ESPECÍFICA
    # ---------------------------------------------------------
    def _buscar_emails_por_data(self, referencia):
        folder = self._get_target_folder()
        mensagens = []

        folder.Items.Sort("[ReceivedTime]", True)

        for msg in folder.Items:
            if msg.Class != 43:  # MailItem
                continue

            if msg.ReceivedTime.date() != referencia.date():
                continue

            assunto = msg.Subject.lower()
            if "relatórios att" not in assunto:
                continue

            for att in msg.Attachments:
                if att.FileName.lower().endswith(".msg"):
                    mensagens.append(msg)
                    break

        return mensagens
    
    # ---------------------------------------------------------
    # BUSCA DO ARQUIVO MAIS ATUAL
    # ---------------------------------------------------------
    def buscar_emails_ultimos_dias(self, dias=3):
        folder = self._get_target_folder()
        mensagens = []

        folder.Items.Sort("[ReceivedTime]", True)  # mais novos primeiro

        hoje = datetime.today().date()

        for msg in folder.Items:
            if msg.Class != 43:
                continue

            assunto = msg.Subject.lower()
            if "relatórios att" not in assunto:
                continue

            # diferença de dias
            diff = (hoje - msg.ReceivedTime.date()).days

            if 0 <= diff < dias:
                # precisa conter um .msg dentro
                for att in msg.Attachments:
                    if att.FileName.lower().endswith(".msg"):
                        mensagens.append(msg)
                        break

        return mensagens

