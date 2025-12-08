import win32com.client as win32
from datetime import datetime, timedelta


class EmailServiceSafra:

    def __init__(self):
        self.outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

    def _get_folder(self):
        inbox = self.outlook.GetDefaultFolder(6)  # Inbox
        itau = inbox.Folders["Itau"]
        safra = itau.Folders["SAFRA"]
        return safra

    def buscar_ultimo_email(self, assunto_alvo: str):
        pasta = self._get_folder()

        mensagens_validas = []
        hoje = datetime.now().date()
        limite = hoje - timedelta(days=3)

        for msg in pasta.Items:
            if msg.Class != 43:
                continue

            assunto = msg.Subject.lower()
            if assunto_alvo.lower() not in assunto:
                continue

            data_email = msg.ReceivedTime.date()

            # Apenas dentro dos últimos 3 dias
            if data_email < limite:
                continue

            mensagens_validas.append(msg)

        if not mensagens_validas:
            return None

        # 🔥 Ordenação correta via Python, ignorando a ordenação bugada do Outlook
        mensagens_validas.sort(key=lambda m: m.ReceivedTime, reverse=True)

        return mensagens_validas[0]  # mais recente
