from pathlib import Path


def salvar_anexo(email, pasta_destino: Path):

    pasta_destino.mkdir(parents=True, exist_ok=True)

    for att in email.Attachments:
        nome = att.FileName

        caminho_final = pasta_destino / nome
        att.SaveAsFile(str(caminho_final))

        return caminho_final

    return None
