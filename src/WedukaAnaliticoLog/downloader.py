import os
import fnmatch


class WedukaDownloader:
    """
    Responsável por salvar o anexo Excel do e-mail
    respeitando o padrão de nome.
    """

    def __init__(self, output_dir):
        self.output_dir = output_dir

    def salvar_anexo_padrao(self, email, padrao_nome, nome_final):
        """
        Localiza o anexo pelo padrão e salva com nome fixo.
        """
        for att in email.Attachments:
            nome_original = att.FileName

            if fnmatch.fnmatch(nome_original, padrao_nome):
                caminho_final = os.path.join(self.output_dir, nome_final)
                att.SaveAsFile(caminho_final)
                return caminho_final

        return None
