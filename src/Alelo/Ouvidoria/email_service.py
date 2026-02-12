import win32com.client as win32
from datetime import datetime, timedelta


class EmailServiceAleloOuvidoria:
    """
    Busca o último e-mail na pasta:
    Caixa de Entrada > Alelo > Intergrall

    Regra:
    - Considera apenas últimos 4 dias
    - Retorna o mais recente encontrado
    """

    def __init__(self):
        self.outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

    def _get_folder(self):
        inbox = self.outlook.GetDefaultFolder(6)  # Caixa de Entrada
        alelo = inbox.Folders["Alelo"]
        intergrall = alelo.Folders["Intergrall"]
        return intergrall

    def buscar_ultimo_email(self, assunto_alvo, dias=4):
        pasta = self._get_folder()
        mensagens = pasta.Items

        # Ordena do mais recente para o mais antigo
        mensagens.Sort("[ReceivedTime]", True)

        hoje = datetime.now().date()
        limite = hoje - timedelta(days=dias)

        for msg in mensagens:
            if msg.Class != 43:
                continue

            # 🔥 Compara apenas DATA (sem timezone)
            data_email = msg.ReceivedTime.date()

            if data_email < limite:
                break

            if assunto_alvo.lower() in (msg.Subject or "").lower():
                return msg

        return None
