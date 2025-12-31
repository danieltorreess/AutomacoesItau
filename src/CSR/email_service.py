import win32com.client as win32
import re
from datetime import datetime, timedelta


class EmailService:
    """
    Serviço responsável por acessar o Outlook, navegar na pasta correta
    e retornar os e-mails do dia de referência contendo a base do discador.
    """

    def __init__(self):
        self.outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
        self.data_referencia = None

    def _get_target_folder(self):
        """
        Navega até a pasta:
        Caixa de Entrada > Itau > CSR
        """
        inbox = self.outlook.GetDefaultFolder(6)  # 6 = Caixa de Entrada
        itau_folder = inbox.Folders["Itau"]
        csr_folder = itau_folder.Folders["CSR"]
        return csr_folder

    def _calcular_data_referencia(self):
        """
        Regra de negócio:
        - Segunda → sexta (-3)
        - Domingo → sexta (-2)
        - Sábado → sexta (-1)
        - Demais → dia anterior
        """
        hoje = datetime.today()
        weekday = hoje.weekday()  # Monday=0

        if weekday == 0:   # Segunda
            return hoje - timedelta(days=3)
        if weekday == 6:   # Domingo
            return hoje - timedelta(days=2)
        if weekday == 5:   # Sábado
            return hoje - timedelta(days=1)

        return hoje - timedelta(days=1)

    def _normalizar_assunto(self, assunto: str) -> str:
        """
        Remove caracteres invisíveis do Outlook (NBSP),
        normaliza espaços e remove lixo de extremidade.
        """
        return (
            assunto
            .replace("\xa0", " ")  # NBSP → espaço normal
            .strip()
        )

    def buscar_emails_do_dia_anterior(self):
        """
        Retorna e-mails no padrão:
        BASE PARA DISCADOR CSR IC DD.MM.YYYY HHMM Horas(.)
        """

        folder = self._get_target_folder()

        referencia = self._calcular_data_referencia()
        self.data_referencia = referencia
        data_str = referencia.strftime("%d.%m.%Y")

        # Regex tolerante a:
        # - múltiplos espaços
        # - ponto final opcional
        regex = rf"BASE PARA DISCADOR CSR IC {data_str}\s+\d{{4}}\s+Horas\.?"

        mensagens_encontradas = []

        for msg in folder.Items:
            if msg.Class != 43:  # MailItem
                continue

            assunto_original = msg.Subject
            assunto = self._normalizar_assunto(assunto_original)

            if re.fullmatch(regex, assunto, re.IGNORECASE):
                mensagens_encontradas.append(msg)

        return mensagens_encontradas

    def get_data_referencia(self):
        return self.data_referencia
