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
python -m pip install browser-cookie3 requests
python -m pip install selenium selenium-wire webdriver-manager
python -m pip install playwright

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

▶️ ItauScout
Move todas as bases do Scout para seus devidos diretórios na rede
python -m src.ItauScout.app

▶️ FalhasOperacionais
Move todas as bases do NGG para seus devidos diretórios na rede
python -m src.FalhasOperacionais.app

▶️ WedukaIncidentes
Navega na ferramenta Weduka, extrai os repositórios listados e move todas as bases de incidentes para seus devidos diretórios na rede
python -m src.WedukaIncidentes.app

▶️ WedukaProcedimentos
Navega na ferramenta Weduka, extrai os repositórios listados e move todas as bases de procedimentos para seus devidos diretórios na rede
python -m src.WedukaProcedimentos.app

▶️ WedukaAnalticoLog
Baixa anexo da base de analítico de log diário do Weduka.
python -m src.WedukaAnaliticoLog.app

▶️ Femme/ReguaAcionamento
Extrai a base da ferramenta live da Femme para régua de acionamento
python -m src.Femme.ReguaAcionamento.app

▶️ SMSLoginLogout
Destrava arquivo excel via XML e monta a base geral de tempos de pausas
python -m src.SMSLoginLogout.app

▶️ OperacaoLibras
Extrair base do anexo e salva na rede
python -m src.OperacaoLibras.app

▶️ FalhasOperacionais
Extrair base do downloads e orquestra para salvar na rede
python -m src.FalhasOperacionais.app

▶️ Envio dos relatórios para o KPI
Envia todos os MIS para o KPI
python -m src.EnvioRelatorios.app

▶️ Férias Alelo
RPAs construídos pela Atento para atualização das bases de Alelo

10. DesenvolvimentoBackEnd/
│
├── downloads/                       # Downloads temporários dos RPAs
│
├── src/
│   │
│   ├── Alelo/
│   │   ├── __pycache__/
│   │   ├── AjustaLayoutBaseBlip.py
│   │   ├── Consolida_Pesquisa.py
│   │   ├── ConverterFormatoAcelera.py
│   │   ├── ETLAleloBKOCredit.py
│   │   ├── IntegrarAleloEdge.py
│   │   └── PlusoftAleloEdge.py
│   │
│   ├── BKO/
│   │   ├── __pycache__/
│   │   ├── app.py
│   │   ├── email_service.py
│   │   └── processor.py
│   │
│   ├── CSR/
│   │   ├── __pycache__/
│   │   ├── __init__.py
│   │   ├── app.py
│   │   ├── downloader.py
│   │   └── email_service.py
│   │
│   ├── EnvioRelatorios/
│   │   ├── __pycache__/
│   │   ├── app.py
│   │   └── email_sender.py
│   │
│   ├── FalhasOperacionais/
│   │   ├── __pycache__/
│   │   └── app.py
│   │
│   ├── Femme/
│   │   └── ReguaAcionamento/
│   │       ├── __pycache__/
│   │       ├── app.py
│   │       ├── browser_edge.py
│   │       ├── config.py
│   │       ├── regua_acionamento_bot.py
│   │       └── utils.py
│   │
│   ├── ItauScout/
│   │   ├── __pycache__/
│   │   └── app.py
│   │
│   ├── ItauSiteGov/
│   │   ├── __pycache__/
│   │   ├── app.py
│   │   ├── browser.py
│   │   ├── config.py
│   │   ├── login.py
│   │   ├── navigation.py
│   │   └── utils.py
│   │
│   ├── OperacaoLibras/
│   │   ├── __pycache__/
│   │   ├── app.py
│   │   ├── downloader.py
│   │   ├── email_service.py
│   │   └── file_utils.py
│   │
│   ├── RAeGOV/
│   │   ├── __pycache__/
│   │   ├── app.py
│   │   ├── email_service.py
│   │   └── processor.py
│   │
│   ├── SAFRA/
│   │   ├── __pycache__/
│   │   ├── app.py
│   │   ├── downloader.py
│   │   ├── email_service.py
│   │   └── excel_utils.py
│   │
│   ├── Shrinkage/
│   │   ├── __pycache__/
│   │   ├── __init__.py
│   │   ├── app.py
│   │   ├── atendimento_processor.py
│   │   ├── downloader.py
│   │   ├── email_service.py
│   │   └── msg_extractor.py
│   │
│   ├── SMSLoginLogout/
│   │   ├── __pycache__/
│   │   ├── app.py
│   │   └── explorador_login_logout_oficial.py
│   │
│   ├── WedukaAnaliticoLog/
│   │   ├── __pycache__/
│   │   ├── app.py
│   │   ├── downloader.py
│   │   ├── email_service.py
│   │   └── file_utils.py
│   │
│   ├── WedukaIncidentes/
│   │   ├── __pycache__/
│   │   ├── app.py
│   │   ├── browser_edge.py
│   │   ├── config.py
│   │   ├── utils.py
│   │   └── weduka_incidentes_bot.py
│   │
│   ├── WedukaProcedimentos/
│   │   ├── __pycache__/
│   │   ├── app.py
│   │   ├── browser_edge.py
│   │   ├── browser.py
│   │   ├── config.py
│   │   ├── utils.py
│   │   ├── weduka_bot.py
│   │   └── README.md
│   │
│   └── __init__.py   # (opcional, mas recomendado)
│
├── tests/
│   ├── debug_position.py
│   ├── testar_explosao.py
│   └── teste.py
│
├── venv/                           # Ambiente virtual (ignorado no Git)
│
├── .env                            # Variáveis de ambiente (ignorado)
├── .env.example
├── .gitignore
├── README.md
├── requirements.txt
├── settings.json
└── settings.example.json

🔐 11. Uso de arquivos de configuração
settings.json → usado localmente
settings.example.json → modelo que deve ir para o GitHub
Nunca enviar credenciais para o repositório.p

🧪 12. Testes auxiliares
python tests/testar_explosao.py
python tests/debug_position.py

🧰 13. Boas práticas
✔ Sempre rodar dentro do venv
✔ Atualizar requirements.txt após instalar novos pacotes
✔ Fazer commits frequentes:

git init
git remote add origin https://github.com/teste/teste.git
git branch -M main
git commit -m "Primeiro commit - Automacoes Itau"
git push -u origin main

Sempre rodar:
git add .
git commit -m "descrição"
git push

✔ Após clonar o repositório rodar as conferências:
git status
git remote -v

✔ Documentar cada nova automação no README




## Alelo
Alelo\ConverterFormatoAcelera.py
\\brsbesrv960\Publico\REPORTS\ALELO\ACELERA