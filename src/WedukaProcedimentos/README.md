1️⃣ VISÃO GERAL — O QUE ESSE PROJETO FAZ

Esse projeto é um RPA (Robotic Process Automation) que:

Abre o Google Chrome

Entra no site do Weduka

Faz login automaticamente

Navega pelos menus

Aplica filtros

Baixa arquivos Excel

Renomeia e move os arquivos

Repete isso para vários repositórios

Tudo isso:

sem API

sem usar mouse/teclado físico

sem travar o computador

de forma automática e confiável

2️⃣ CONCEITOS BÁSICOS DE PYTHON (ANTES DO CÓDIGO)

Antes de entrar nos arquivos, vamos alinhar conceitos.

📦 O que são bibliotecas (imports)

Em Python, biblioteca é código pronto que outra pessoa escreveu para você usar.

Exemplo:

import time


Isso significa:

“Python, me dá acesso às funções de tempo (sleep, timestamp etc).”

No nosso projeto usamos:

bibliotecas padrão do Python

bibliotecas externas (selenium, dotenv)

🧠 O que é uma função

Função = bloco de código reutilizável.

def soma(a, b):
    return a + b


Você chama assim:

resultado = soma(2, 3)


Por que usar função?

evita repetição

organiza o código

facilita manutenção

🏛️ O que é uma classe

Classe = um “molde” de comportamento.

Ela agrupa:

dados

funções relacionadas

Exemplo simples:

class Carro:
    def acelerar(self):
        print("Acelerando")

🔑 O que é self

Esse é um dos pontos mais importantes.

👉 self representa a própria instância da classe.

Quando você escreve:

self.driver


Você está dizendo:

“O driver que pertence a ESTE objeto”

Sem self, a classe não consegue guardar estado.

🧩 O que é um método

Método = função dentro de uma classe.

class Exemplo:
    def metodo(self):
        pass


Toda função dentro de classe recebe self

Métodos operam sobre os dados da classe

🧪 O que é lambda (usamos uma vez)

Lambda é uma função anônima, de uma linha só.

lambda f: f.stat().st_mtime


Significa:

“Recebe f e retorna f.stat().st_mtime”

Usado quando:

a função é simples

não vale criar def

3️⃣ ARQUITETURA DO PROJETO (ORQUESTRAÇÃO)
app.py  → ponto de entrada
  ↓
browser.py → cria o Chrome
  ↓
weduka_bot.py → executa o RPA
  ↓
utils.py → funções auxiliares
  ↓
config.py → parâmetros fixos


Cada arquivo tem uma responsabilidade clara.

Isso é boa prática profissional.

4️⃣ EXPLICAÇÃO ARQUIVO POR ARQUIVO (LINHA A LINHA)
📁 browser.py
Objetivo

👉 Criar e configurar o navegador Chrome.

from selenium import webdriver


Importa o Selenium, que controla o navegador.

from selenium.webdriver.chrome.options import Options


Permite configurar o Chrome (downloads, popups etc).

def get_browser(download_dir: str):


Define uma função que cria o navegador.

download_dir = pasta onde os arquivos serão baixados

: str é tipagem (opcional), só para clareza

options = Options()


Cria o objeto de configuração do Chrome.

options.add_argument("--disable-notifications")


Desliga notificações do Chrome.

Se não fizer isso:
❌ popup pode travar o RPA

prefs = {
    "download.default_directory": str(download_dir),


Define a pasta padrão de download.

"download.prompt_for_download": False


Evita aquela pergunta:

“Deseja salvar este arquivo?”

options.add_experimental_option("prefs", prefs)


Aplica essas configurações no Chrome.

driver = webdriver.Chrome(options=options)


Cria o Chrome usando o Selenium Manager (automático).

return driver


Devolve o navegador para quem chamou a função.

📁 config.py
Objetivo

👉 Centralizar configurações.

from pathlib import Path


Biblioteca moderna para trabalhar com caminhos de arquivos.

URL_INTEGRATION = "https://..."


URL inicial do sistema.

Se mudar no futuro:

altera só aqui

DOWNLOAD_DIR = Path.home() / "Downloads"


Pasta padrão de downloads do usuário.

DEST_DIR = Path(r"\\SERVIDOR\PASTA")


Pasta final onde o SSIS vai ler os arquivos.

REPOSITORIOS = [
    "Cartões",
    "Cartões PJ",
]


Lista de repositórios que o RPA vai processar.

👉 Isso é programação orientada a dados.

FILE_PREFIX = "download_procedimentos_"


Prefixo usado para padronizar os arquivos.

📁 utils.py
Objetivo

👉 Funções auxiliares reutilizáveis.

from datetime import datetime


Usado para trabalhar com datas.

def get_date_range():


Função que gera:

01/MM/YYYY - hoje

hoje = datetime.today()


Pega a data atual do sistema.

inicio = hoje.replace(day=1)


Força o dia para 01.

strftime("%d/%m/%Y")


Formata a data no padrão brasileiro.

def wait_for_download(download_dir: Path, timeout=120):


Função que espera o download terminar.

timeout = tempo máximo de espera

files = list(download_dir.glob("download_*.xlsx"))


Procura arquivos que começam com download_.

max(files, key=lambda f: f.stat().st_mtime)


Escolhe o arquivo mais recente.

def move_and_rename(file, dest_dir, new_name):


Move e renomeia o arquivo baixado.

shutil.move(...)


Move o arquivo de forma segura.

📁 weduka_bot.py
Objetivo

👉 Executar o fluxo do site.

class WedukaBot:


Define uma classe que representa o robô.

def __init__(self, driver, username, password, config):


Construtor da classe.

Executa quando você faz:

bot = WedukaBot(...)

self.driver = driver


Guarda o navegador dentro do objeto.

self.wait = WebDriverWait(driver, 30)


Define espera inteligente:

espera até 30s

evita sleep fixo

🔐 login()
self.driver.get(self.config.URL_INTEGRATION)


Abre o site.

(By.LINK_TEXT, "Ir para site de autenticação")


Localiza o botão pelo texto visível.

send_keys(self.username)


Digita o usuário.

📊 acessar_relatorio()

Navega pelos menus usando:

texto visível

espera inteligente

Isso é mais estável que ID dinâmico.

🧠 extrair_repositorio()

Aqui está o coração do RPA.

Dropdown Bootstrap
.dropdown-menu.show


Garante que o dropdown está aberto.

safe_click

Função criada para:

evitar overlay

evitar erro de clique interceptado

Data
date_input.send_keys(Keys.ENTER)


Força o site a aplicar o filtro.

Exportação (parte mais importante)
export_url = export_link.get_attribute("href")
self.driver.get(export_url)


👉 Não clicamos no link
👉 Navegamos direto para a URL

Isso ignora:

UI

overlay

layout

scroll

É a técnica mais profissional de RPA web.

📁 app.py
Objetivo

👉 Orquestrar tudo.

load_dotenv(env_path)


Carrega usuário e senha sem deixar no código.

driver = get_browser(...)


Cria o Chrome.

bot = WedukaBot(...)


Cria o robô.

for repo in config.REPOSITORIOS:


Loop que permite escalar facilmente.

5️⃣ COMO TUDO SE CONECTA (ORQUESTRAÇÃO FINAL)

app.py inicia

.env é carregado

Chrome é criado

Bot é instanciado

Login é feito

Menu é acessado

Loop de repositórios começa

Arquivo é baixado

Arquivo é renomeado

SSIS consome