✅ README.md — Automações Itaú
🏦 Automações Itaú

Repositório contendo diversas automações voltadas ao processamento de bases recebidas via Outlook, integrando com Excel, CSVs e diretórios da rede interna.

🔧 1. Requisitos

Python 3.10+ instalado

Permissão para rodar scripts Python no SISGAG

Microsoft Outlook instalado e com conta autenticada

Acesso aos diretórios da rede especificados nas automações

🐍 2. Verificando a instalação do Python
python --version

Se o comando não funcionar, tente:
py --version

🌱 3. Criando o ambiente virtual
python -m venv venv

▶️ 4. Ativando o ambiente virtual
.\venv\Scripts\activate

Para desativar:
deactivate

🚫 5. Importante no SISGAG (restrição de segurança)

Nunca usar pip install diretamente.
O SISGAG bloqueia. Sempre usar python -m pip.

Exemplos corretos:

python -m pip install --upgrade pip
python -m pip install pandas openpyxl
python -m pip install pywin32

Exemplo incorreto (bloqueado pelo SISGAG):
pip install pandas

📦 6. Listando pacotes instalados
python -m pip list
python -m pip freeze

📄 7. Gerando requirements.txt
python -m pip freeze > requirements.txt

📥 8. Instalando pacotes a partir do requirements
python -m pip install -r requirements.txt

🚀 9. Executando cada automação
▶️ Shrinkage
Processa as dinâmicas VOZ e DIGITAL do arquivo ATT.
python -m src.Shrinkage.app

▶️ SAFRA
Baixa e processa arquivos dos e-mails GERENCIAL_LOG e MIS31047.
python -m src.SAFRA.app

▶️ RAeGOV
Baixa arquivos do Consignado e ajusta o nome das abas.
python -m src.RAeGOV.app

▶️ BKO
Busca bases do dia atual, limpa ou salva CSVs e envia para a rede.
python -m src.BKO.app

▶️ CSR
Baixa e processa arquivos dos e-mails
python -m src.CSR.app

🗂 10. Estrutura completa do projeto
DesenvolvimentoBackEnd/
│
├── src/
│   ├── BKO/
│   │   ├── app.py
│   │   ├── email_service.py
│   │   ├── processor.py
│   │
│   ├── RAeGOV/
│   │   ├── app.py
│   │   ├── email_service.py
│   │   ├── processor.py
│   │
│   ├── SAFRA/
│   │   ├── app.py
│   │   ├── downloader.py
│   │   ├── email_service.py
│   │   ├── excel_utils.py
│   │
│   ├── Shrinkage/
│   │   ├── app.py
│   │   ├── downloader.py
│   │   ├── email_service.py
│   │   ├── atendimento_processor.py
│   │   ├── msg_extractor.py
│
├── tests/
│   ├── debug_position.py
│   ├── testar_explosao.py
│   ├── teste.py
│
├── venv/                     # Ignorado no Git
│
├── .gitignore
├── README.md
├── requirements.txt
├── settings.json
└── settings.example.json

🔐 11. Uso de arquivos de configuração
settings.json → usado localmente
settings.example.json → modelo que deve ir para o GitHub
Nunca enviar credenciais para o repositório.

🧪 12. Testes auxiliares
python tests/testar_explosao.py
python tests/debug_position.py

🧰 13. Boas práticas
✔ Sempre rodar dentro do venv
✔ Atualizar requirements.txt após instalar novos pacotes
✔ Fazer commits frequentes:

git init
git remote add origin https://github.com/danieltorreess/AutomacoesItau.git
git branch -M main
git commit -m "Primeiro commit - Automacoes Itau"
git push -u origin main

Sempre rodar:
git add .
git commit -m "descrição"
git push

✔ Documentar cada nova automação no README