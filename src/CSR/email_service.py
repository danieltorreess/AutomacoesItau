import win32com.client as win32
import re
from datetime import datetime, timedelta


class EmailService:
    """
    Serviço responsável por acessar o Outlook, navegar na pasta correta
    e retornar os e-mails do dia anterior contendo a base do discador.
    """

    def __init__(self):
        # Conecta no Outlook instalado
        self.outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

    def _get_target_folder(self):
        """
        Navega até a pasta:
        Caixa de Entrada > Itau > CSR
        """
        inbox = self.outlook.GetDefaultFolder(6)  # 6 = Caixa de entrada

        # Pasta > Subpasta
        itau_folder = inbox.Folders["Itau"]
        csr_folder = itau_folder.Folders["CSR"]

        return csr_folder

    def buscar_emails_do_dia_anterior(self):
        """
        Retorna os e-mails do padrão:
        'BASE PARA DISCADOR CSR IC DD.MM.YYYY HHMM Horas'
        referentes ao dia anterior.
        """

        folder = self._get_target_folder()

        # Data de referência: dia anterior
        referencia = datetime.today() - timedelta(days=1)
        data_str = referencia.strftime("%d.%m.%Y")

        # Regex para capturar o padrão
        regex = rf"BASE PARA DISCADOR CSR IC {data_str} \d{{4}} Horas"

        mensagens_encontradas = []

        # Percorre os e-mails da pasta
        for msg in folder.Items:
            if msg.Class != 43:  # 43 = MailItem
                continue

            assunto = msg.Subject.strip()

            # Verifica se bate com o padrão
            if re.fullmatch(regex, assunto):
                mensagens_encontradas.append(msg)

        return mensagens_encontradas
