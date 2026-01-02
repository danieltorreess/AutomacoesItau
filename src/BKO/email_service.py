# src/BKO/email_service.py
import win32com.client as win32
from datetime import datetime
import unicodedata
import re

class EmailServiceBKO:

    def __init__(self):
        outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
        self.folder = (
            outlook.GetDefaultFolder(6)  # inbox
            .Folders["Itau"]
            .Folders["BKO"]
        )

    # -------------------------------------------------------

    def _is_today(self, msg):
        return msg.ReceivedTime.date() == datetime.today().date()

    def _normalizar(self, texto: str) -> str:
        texto = texto.lower()
        texto = unicodedata.normalize("NFKD", texto)
        texto = "".join(c for c in texto if not unicodedata.combining(c))
        texto = re.sub(r"\s+", " ", texto)
        return texto.strip()

    # -------------------------------------------------------

    def buscar_email_pos_venda(self):
        """
        Busca e-mail do DIA ATUAL relacionado a Pós Venda,
        independente de acento ou formato.
        """
        for msg in sorted(self.folder.Items, key=lambda x: x.ReceivedTime, reverse=True):
            if not self._is_today(msg):
                continue

            subj = self._normalizar(msg.Subject)

            if "pos venda" in subj:
                return msg

        return None

    # -------------------------------------------------------

    def buscar_email_segundo_nivel(self):
        """
        Busca e-mail do DIA ATUAL contendo 'Base 2° nível atendimento PJ'.
        Suporta variações: "2º", "2°", "2o", etc.
        """
        patterns = ["base 2", "nível atendimento pj", "nivel atendimento pj"]

        for msg in sorted(self.folder.Items, key=lambda x: x.ReceivedTime, reverse=True):
            if not self._is_today(msg):
                continue
            subj = msg.Subject.lower()

            if all(p in subj for p in ["base", "pj"]):
                if any(p in subj for p in patterns):
                    return msg

        return None
