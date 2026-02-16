import os


class ChargeBackDownloader:

    def __init__(self, pasta_destino):
        self.pasta_destino = pasta_destino
        os.makedirs(self.pasta_destino, exist_ok=True)

    def salvar_todos_anexos(self, email):

        anexos_salvos = []

        print("📎 Verificando anexos...")

        for att in email.Attachments:

            nome = att.FileName

            if nome.lower().endswith(".xlsx"):

                caminho = os.path.join(self.pasta_destino, nome)
                att.SaveAsFile(caminho)

                print(f"📥 Anexo salvo: {nome}")
                anexos_salvos.append(caminho)

        return anexos_salvos
