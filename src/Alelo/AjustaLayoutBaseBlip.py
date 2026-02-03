
import os
import zipfile
import pandas as pd

# Caminho da pasta onde estão os arquivos
base_path = r"\\brsbesrv960\publico\REPORTS\ALELO\BLIP"

# Prefixo do arquivo
file_prefix = "AgentHistory_prdtransbordoecs_"

# Lista de colunas na ordem desejada
ordered_columns = [
    "SequentialId","CustomerIdentity","AgentIdentity","Status","StorageDate","ExpirationDate","CloseDate","Team","Closed","Tags","QueueTime","FirstResponseTime","AverageResponseTime","CustomerName","CustomerEmail","CustomerGender","CustomerCity","CustomerPhoneNumber","AgentName","AgentEmail","tunnel.owner","tunnel.originator","LGPD","CNPJ","Protocolo","Optin","Fila","Segmento","credentialingStep1","credentialingStep2","isLoggedIn","EC","NomeEC","credentialingStep3","protocoloVenda","botIterationsCounter","media","activeMessageFileName","campaignId","campaignMessageTemplate","campaignOriginator"
]

# 1. Localiza o arquivo ZIP
zip_file = None
for file in os.listdir(base_path):
    if file.startswith(file_prefix) and file.endswith(".zip"):
        zip_file = os.path.join(base_path, file)
        break

if not zip_file:
    raise FileNotFoundError("Nenhum arquivo .zip encontrado com o prefixo especificado.")

# 2. Extrai o arquivo ZIP
with zipfile.ZipFile(zip_file, 'r') as zip_ref:
    zip_ref.extractall(base_path)

# 3. Procura o arquivo CSV com o mesmo prefixo
csv_file = None
for file in os.listdir(base_path):
    if file.startswith(file_prefix) and file.endswith(".csv"):
        csv_file = os.path.join(base_path, file)
        break

if not csv_file:
    raise FileNotFoundError(f"Nenhum arquivo CSV encontrado com prefixo {file_prefix}.")

# 4. Detecta delimitador automaticamente
with open(csv_file, 'r', encoding='utf-8') as f:
    first_line = f.readline()
    delimiter = ';' if first_line.count(';') > first_line.count(',') else ','

# 5. Lê o arquivo CSV com o delimitador correto
df = pd.read_csv(csv_file, sep=delimiter, dtype=str, on_bad_lines='skip')

# 6. Mantém apenas as colunas desejadas e na ordem especificada
existing_columns = [col for col in ordered_columns if col in df.columns]
df = df[existing_columns]

# 7. Sobrescreve o arquivo original com colunas separadas
df.to_csv(csv_file, index=False, encoding='utf-8', sep=delimiter)

print(f"✅ Processo concluído com sucesso! Arquivo atualizado: {csv_file}")
print(f"Delimitador detectado: {delimiter}")

# 8. Exclui o arquivo ZIP ao final do processo
try:
    os.remove(zip_file)
    print(f"🗑️ Arquivo ZIP removido: {zip_file}")
except Exception as e:
    print(f"⚠️ Não foi possível remover o ZIP ({zip_file}). Erro: {e}")
