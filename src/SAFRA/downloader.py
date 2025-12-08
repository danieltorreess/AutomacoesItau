import os


class SafraDownloader:

    def __init__(self, output_dir):
        self.output_dir = output_dir

    def salvar_anexo_excel(self, email):
        """
        Salva o primeiro anexo Excel do e-mail.
        Sempre sobrescreve.
        """
        for att in email.Attachments:
            nome = att.FileName.lower()

            if nome.endswith(".xlsx") or nome.endswith(".xlsm") or nome.endswith(".xls"):
                caminho = os.path.join(self.output_dir, att.FileName)
                att.SaveAsFile(caminho)
                return caminho

        return None
