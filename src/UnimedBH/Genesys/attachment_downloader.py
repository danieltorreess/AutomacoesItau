import os

from .config import ATTACHMENTS, BASE_PATH
from .prestador_parser import tratar_prestador
from .backup_service import mover_para_bkp
from .logger import log


class AttachmentDownloader:

    def salvar_anexos(self, email):

        arquivos = []

        for att in email.Attachments:

            nome = att.FileName

            for tipo, nome_padrao in ATTACHMENTS.items():

                if nome == nome_padrao:

                    caminho = os.path.join(BASE_PATH, nome)

                    # mover arquivo existente para BKP
                    if os.path.exists(caminho):

                        log(f"📦 Arquivo existente encontrado: {nome}")

                        mover_para_bkp(tipo, caminho)

                    # salvar novo
                    att.SaveAsFile(caminho)

                    log(f"💾 Novo arquivo salvo na rede: {caminho}")

                    # tratar prestador
                    if tipo == "PRESTADOR":

                        log("🛠 Tratando arquivo PRESTADOR")

                        tratar_prestador(caminho)

                    arquivos.append((tipo, caminho))

        return arquivos