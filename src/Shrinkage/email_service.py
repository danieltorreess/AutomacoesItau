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

        agora = datetime.now().replace(tzinfo=None)
        limite = (agora - timedelta(days=dias)).replace(tzinfo=None)

        items = folder.Items
        items.Sort("[ReceivedTime]", True)

        for msg in items:
            if msg.Class != 43:
                continue

            recebido = msg.ReceivedTime.replace(tzinfo=None)

            if recebido < limite:
                break

            assunto = (msg.Subject or "").lower()

            if "relatórios att" not in assunto:
                continue

            for att in msg.Attachments:
                if att.FileName.lower().endswith((".xls", ".xlsx", ".xlsm", ".xlsb")):
                    mensagens.append(msg)
                    break

        return mensagens


