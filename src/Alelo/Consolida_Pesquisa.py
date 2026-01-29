#type:ignore
import glob
import shutil
import paramiko
import os
from datetime import datetime

# Credenciais e caminhos
host = "10.235.235.158"
port = 22
username = "631227"
password = "Lilo!@220886"
caminho_remoto = "/home_portal/631227/SFTP_ALELO_TLV/CSAT_DNK"
caminho_local = r"C:\Users\ab1541240\Downloads\SFTP_PESQUISA"

# Função para extrair timestamp do nome do arquivo
def extrair_timestamp(nome):
    try:
        parte = nome.split("BI_IBK_PESQUISA_")[1].split(".txt")[0]
        return datetime.strptime(parte, "%Y%m%d%H%M")
    except:
        return None

# Verifica o último arquivo local
arquivos_locais = [f for f in os.listdir(caminho_local) if f.startswith("BI_IBK_PESQUISA_") and f.endswith(".txt")]
if arquivos_locais:
    ultimo_arquivo = max(arquivos_locais, key=extrair_timestamp)
    timestamp_ultimo = extrair_timestamp(ultimo_arquivo)
else:
    timestamp_ultimo = datetime.min

# Conecta à SFTP e lista arquivos
transport = paramiko.Transport((host, port))
transport.connect(username=username, password=password)
sftp = paramiko.SFTPClient.from_transport(transport)

arquivos_remotos = sftp.listdir(caminho_remoto)

# Filtra e baixa arquivos mais recentes
baixados = 0
for arquivo in sorted(arquivos_remotos):
    if arquivo.startswith("BI_IBK_PESQUISA_") and arquivo.endswith(".txt"):
        timestamp_arquivo = extrair_timestamp(arquivo)
        if timestamp_arquivo and timestamp_arquivo > timestamp_ultimo:
            caminho_completo_remoto = caminho_remoto + "/" + arquivo
            caminho_completo_local = os.path.join(caminho_local, arquivo)
            sftp.get(caminho_completo_remoto, caminho_completo_local)
            baixados += 1

sftp.close()
transport.close()

print(f"{baixados} arquivos foram baixados da SFTP para {caminho_local}.")

# Caminhos das pastas
pasta_origem = r"C:\Users\ab1541240\Downloads\SFTP_PESQUISA"
pasta_destino = r"C:\Users\ab1541240\Downloads\SFTP_PESQUISA\CONSOLIDADO"

# Garante que a pasta de destino existe
os.makedirs(pasta_destino, exist_ok=True)

# Nome do arquivo consolidado
arquivo_saida = os.path.join(pasta_destino, "BI_IBK_PESQUISA_ALELO.txt")

# Padrão dos arquivos a serem consolidados
padrao_arquivos = os.path.join(pasta_origem, "BI_IBK_PESQUISA_*.txt")

# Lista e ordena os arquivos
arquivos_txt = sorted(glob.glob(padrao_arquivos))

# Inicializa variável para armazenar o cabeçalho
cabecalho = None

# Consolida os arquivos
with open(arquivo_saida, 'w', encoding='utf-8') as consolidado:
    for idx, arquivo in enumerate(arquivos_txt):
        try:
            with open(arquivo, 'r', encoding='latin-1') as f:
                linhas = f.readlines()

                # Se for o primeiro arquivo, salva o cabeçalho
                if idx == 0:
                    cabecalho = linhas[0].strip()
                    consolidado.write(cabecalho + "\n")

                # Escreve os dados (ignorando o cabeçalho dos demais arquivos)
                for linha in linhas[1:]:
                    linha_limpa = linha.strip()
                    if linha_limpa:
                        consolidado.write(linha_limpa + "\n")
        except Exception as e:
            print(f"Erro ao ler o arquivo {arquivo}: {e}")

print(f"{len(arquivos_txt)} arquivos foram processados. Consolidado salvo em: {arquivo_saida}")


# Caminho do arquivo consolidado
arquivo_consolidado = r"C:\Users\ab1541240\Downloads\SFTP_PESQUISA\CONSOLIDADO\BI_IBK_PESQUISA_ALELO.txt"

# Caminho de rede de destino
destino_rede = r"\\brsbesrv960\publico\REPORTS\ALELO\CSAT\BI_IBK_PESQUISA_ALELO.txt"

# Tenta copiar o arquivo
try:
    shutil.copy2(arquivo_consolidado, destino_rede)
    print(f"Arquivo copiado com sucesso para: {destino_rede}")
except Exception as e:
    print(f"Erro ao copiar o arquivo: {e}")
