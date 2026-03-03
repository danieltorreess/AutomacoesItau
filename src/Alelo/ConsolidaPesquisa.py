# type: ignore
import glob
import shutil
import paramiko
from pathlib import Path
from datetime import datetime
from scp import SCPClient


# ==============================
# CONFIGURAÇÕES (TUDO AQUI)
# ==============================

HOST = "10.235.235.158"
PORT = 22
USERNAME = "631227"
PASSWORD = "Lilo!@220886" 

CAMINHO_REMOTO = "/home_portal/631227/SFTP_ALELO_TLV/CSAT_DNK"

CAMINHO_LOCAL = Path(r"C:\Users\ab1541240\Downloads\SFTP_PESQUISA")
PASTA_CONSOLIDADO = CAMINHO_LOCAL / "CONSOLIDADO"
DESTINO_REDE = Path(r"\\brsbesrv960\publico\REPORTS\ALELO\CSAT\BI_IBK_PESQUISA_ALELO.txt")

PADRAO_ARQUIVO = "BI_IBK_PESQUISA_"
ARQUIVO_FINAL = PASTA_CONSOLIDADO / "BI_IBK_PESQUISA_ALELO.txt"


# ==============================
# FUNÇÕES
# ==============================

def extrair_timestamp(nome):
    try:
        parte = nome.split(PADRAO_ARQUIVO)[1].split(".txt")[0]
        return datetime.strptime(parte, "%Y%m%d%H%M")
    except:
        return None

def conectar_sftp():
    print("🔐 Conectando à SFTP...")

    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())

    ssh.connect(
        hostname=HOST,
        port=PORT,
        username=USERNAME,
        password=PASSWORD,
        timeout=15
    )

    print("✅ SSH conectado")

    # 🔎 TESTE DE COMANDO REMOTO
    stdin, stdout, stderr = ssh.exec_command("pwd")
    print("📂 Diretório atual:", stdout.read().decode())

    stdin, stdout, stderr = ssh.exec_command(
        "ls /home_portal/631227/SFTP_ALELO_TLV/CSAT_DNK"
    )
    print("📁 Arquivos no diretório alvo:")
    print(stdout.read().decode())

    ssh.close()
    exit()

def obter_ultimo_timestamp_local():
    CAMINHO_LOCAL.mkdir(parents=True, exist_ok=True)

    arquivos = [
        f.name for f in CAMINHO_LOCAL.glob(f"{PADRAO_ARQUIVO}*.txt")
    ]

    if not arquivos:
        return datetime.min

    ultimo = max(arquivos, key=extrair_timestamp)
    return extrair_timestamp(ultimo) or datetime.min


def baixar_arquivos_scp(transport):

    print("📂 Listando arquivos via SSH...")

    ssh = paramiko.SSHClient()
    ssh._transport = transport

    stdin, stdout, stderr = ssh.exec_command(
        f"ls {CAMINHO_REMOTO}"
    )

    arquivos_remotos = stdout.read().decode().splitlines()

    timestamp_ultimo = obter_ultimo_timestamp_local()
    print(f"📅 Último timestamp local: {timestamp_ultimo}")

    baixados = 0

    with SCPClient(transport) as scp:

        for arquivo in sorted(arquivos_remotos):

            if arquivo.startswith(PADRAO_ARQUIVO) and arquivo.endswith(".txt"):

                timestamp = extrair_timestamp(arquivo)

                if timestamp and timestamp > timestamp_ultimo:

                    remoto = f"{CAMINHO_REMOTO}/{arquivo}"
                    local = CAMINHO_LOCAL / arquivo

                    print(f"⬇ Baixando via SCP: {arquivo}")
                    scp.get(remoto, str(local))
                    baixados += 1

    print(f"📥 {baixados} arquivo(s) baixado(s).")


def consolidar_arquivos():
    print("🧩 Consolidando arquivos...")

    PASTA_CONSOLIDADO.mkdir(parents=True, exist_ok=True)

    arquivos_txt = sorted(
        glob.glob(str(CAMINHO_LOCAL / f"{PADRAO_ARQUIVO}*.txt"))
    )

    if not arquivos_txt:
        print("⚠ Nenhum arquivo para consolidar.")
        return 0

    with open(ARQUIVO_FINAL, "w", encoding="utf-8") as consolidado:

        for idx, arquivo in enumerate(arquivos_txt):
            try:
                with open(arquivo, "r", encoding="latin-1") as f:
                    linhas = f.readlines()

                    if idx == 0:
                        consolidado.write(linhas[0].strip() + "\n")

                    for linha in linhas[1:]:
                        linha = linha.strip()
                        if linha:
                            consolidado.write(linha + "\n")

            except Exception as e:
                print(f"❌ Erro ao processar {arquivo}: {e}")

    print(f"✅ Consolidado gerado: {ARQUIVO_FINAL}")
    return len(arquivos_txt)


def copiar_para_rede():
    print("📡 Copiando para rede...")

    DESTINO_REDE.parent.mkdir(parents=True, exist_ok=True)

    shutil.copy2(ARQUIVO_FINAL, DESTINO_REDE)

    print(f"✅ Arquivo copiado para: {DESTINO_REDE}")


# ==============================
# MAIN
# ==============================

def main():

    print("\n🚀 Iniciando Consolidação Pesquisa (SCP)")
    print("=" * 60)

    transport = None

    try:
        print("🔐 Conectando (modo compatível)...")

        transport = paramiko.Transport((HOST, PORT))

        security_options = transport.get_security_options()
        security_options.key_types = ["ssh-rsa"]
        security_options.kex = [
            "diffie-hellman-group14-sha1",
            "diffie-hellman-group1-sha1"
        ]

        transport.connect(username=USERNAME, password=PASSWORD)

        print("✅ SSH conectado")

        baixar_arquivos_scp(transport)

    except Exception as e:
        print(f"❌ Erro na conexão: {e}")
        return

    finally:
        if transport:
            transport.close()
        print("🔒 Conexão encerrada.")

    total = consolidar_arquivos()

    if total > 0:
        try:
            copiar_para_rede()
        except Exception as e:
            print(f"❌ Erro ao copiar para rede: {e}")

    print("🏁 Processo finalizado.\n")

if __name__ == "__main__":
    main()