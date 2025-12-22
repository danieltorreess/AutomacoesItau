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
        self.data_referencia = None

    def _get_target_folder(self):
        """
        Navega até a pasta:
        Caixa de Entrada > Itau > CSR
        """
        inbox = self.outlook.GetDefaultFolder(6)  # 6 = Caixa de entrada
        itau_folder = inbox.Folders["Itau"]
        csr_folder = itau_folder.Folders["CSR"]
        return csr_folder

    def _calcular_data_referencia(self):
        """
        Regra de negócio para definição da data:
        - Segunda-feira → sexta-feira (-3 dias)
        - Domingo → sexta-feira (-2 dias)
        - Sábado → sexta-feira (-1 dia)
        - Demais dias → dia anterior
        """
        hoje = datetime.today()
        weekday = hoje.weekday()  # Monday=0, Sunday=6

        if weekday == 0:  # Segunda
            return hoje - timedelta(days=3)

        if weekday == 6:  # Domingo
            return hoje - timedelta(days=2)

        if weekday == 5:  # Sábado
            return hoje - timedelta(days=1)

        return hoje - timedelta(days=1)

    def buscar_emails_do_dia_anterior(self):
        """
        Retorna os e-mails do padrão:
        'BASE PARA DISCADOR CSR IC DD.MM.YYYY HHMM Horas'
        referentes à data de referência calculada.
        """

        folder = self._get_target_folder()

        # 🔥 AGORA A DATA VEM DA REGRA CORRETA
        referencia = self._calcular_data_referencia()
        self.data_referencia = referencia
        data_str = referencia.strftime("%d.%m.%Y")

        regex = rf"BASE PARA DISCADOR CSR IC {data_str} \d{{4}} Horas"

        mensagens_encontradas = []

        for msg in folder.Items:
            if msg.Class != 43:  # 43 = MailItem
                continue

            assunto = msg.Subject.strip()

            if re.fullmatch(regex, assunto):
                mensagens_encontradas.append(msg)

        return mensagens_encontradas
    
    def get_data_referencia(self):
        return self.data_referencia
