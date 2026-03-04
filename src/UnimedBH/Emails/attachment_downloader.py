import os
from datetime import datetime

from .config import (
    ATTACHMENT_AGENDAMENTO,
    ATTACHMENT_AUTORIZACAO,
    PASTA_AGENDAMENTO,
    PASTA_AUTORIZACAO,
)


class AttachmentDownloader:

    def salvar_anexos(self, email):

        data_email = email.ReceivedTime.strftime("%d.%m.%y")

        arquivos_salvos = []

        for att in email.Attachments:

            nome = att.FileName

            if nome == ATTACHMENT_AGENDAMENTO:

                novo_nome = f"DESEMPENHO DA FILA - AGENDAMENTO ATENTO D1_{data_email}.csv"

                caminho = os.path.join(PASTA_AGENDAMENTO, novo_nome)

                att.SaveAsFile(caminho)

                arquivos_salvos.append(caminho)

            elif nome == ATTACHMENT_AUTORIZACAO:

                novo_nome = f"DESEMPENHO DA FILA - AUTORIZACAO D1_{data_email}.csv"

                caminho = os.path.join(PASTA_AUTORIZACAO, novo_nome)

                att.SaveAsFile(caminho)

                arquivos_salvos.append(caminho)

        return arquivos_salvos