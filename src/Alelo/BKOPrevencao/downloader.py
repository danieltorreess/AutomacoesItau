from pathlib import Path


def salvar_anexo(email, pasta_destino: Path):

    # 🔥 GARANTIR que a pasta exista antes do SaveAsFile
    pasta_destino.mkdir(parents=True, exist_ok=True)

    for att in email.Attachments:
        if att.FileName.lower().endswith(".xlsx"):

            caminho_final = pasta_destino / att.FileName

            print(f"📎 Salvando anexo: {att.FileName}")
            print(f"📂 Caminho destino: {caminho_final.resolve()}")

            att.SaveAsFile(str(caminho_final.resolve()))

            return caminho_final

    return None