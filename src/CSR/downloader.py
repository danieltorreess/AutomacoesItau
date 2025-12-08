import os

def salvar_anexos(emails, output_path):
    """
    Salva somente anexos .xlsx dos e-mails recebidos.
    Ignora imagens, assinaturas e outros tipos de arquivos.
    """

    # Garante que o diretório existe
    os.makedirs(output_path, exist_ok=True)

    for msg in emails:
        if msg.Attachments.Count == 0:
            print(f"⚠️ Email sem anexos: {msg.Subject}")
            continue

        for attachment in msg.Attachments:
            nome_arquivo = attachment.FileName

            # Ignora anexos sem nome ou sem extensão
            if not nome_arquivo or "." not in nome_arquivo:
                print(f"⏭️ Ignorado (sem nome): {nome_arquivo}")
                continue

            # Aceita apenas arquivos .xlsx
            if not nome_arquivo.lower().endswith(".xlsx"):
                print(f"⏭️ Ignorado (não é Excel): {nome_arquivo}")
                continue

            destino = os.path.join(output_path, nome_arquivo)

            try:
                attachment.SaveAsFile(destino)
                print(f"📁 Arquivo salvo: {destino}")
            except Exception as e:
                print(f"❌ Erro ao salvar {nome_arquivo}: {e}")
