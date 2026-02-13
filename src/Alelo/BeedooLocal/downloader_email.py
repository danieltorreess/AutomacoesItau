from pathlib import Path
import fnmatch


def salvar_anexo_user(email, pasta_destino: Path):

    pasta_destino.mkdir(parents=True, exist_ok=True)

    for att in email.Attachments:
        nome = att.FileName

        if fnmatch.fnmatch(nome, "user*.xlsx"):
            caminho_final = pasta_destino / nome
            att.SaveAsFile(str(caminho_final))
            return caminho_final

    return None
