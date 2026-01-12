import win32com.client as win32
from datetime import datetime, timedelta


class EmailServiceWeduka:
    """
    Responsável por acessar o Outlook e buscar o último e-mail
    da pasta: Caixa de Entrada > Itau > WedukaTreinamentos
    """

    def __init__(self):
        self.outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

    def _get_folder(self):
        inbox = self.outlook.GetDefaultFolder(6)  # Caixa de Entrada
        itau = inbox.Folders["Itau"]
        weduka = itau.Folders["WedukaTreinamentos"]
        return weduka

    def buscar_ultimo_email(self, assunto_alvo: str, dias=5):
        pasta = self._get_folder()

        mensagens_validas = []
        hoje = datetime.now().date()
        limite = hoje - timedelta(days=dias)

        for msg in pasta.Items:
            if msg.Class != 43:  # MailItem
                continue

            if assunto_alvo.lower() not in msg.Subject.lower():
                continue

            if msg.ReceivedTime.date() < limite:
                continue

            mensagens_validas.append(msg)

        if not mensagens_validas:
            return None

        mensagens_validas.sort(key=lambda m: m.ReceivedTime, reverse=True)
        return mensagens_validas[0]
