# type:ignore
# -*- coding: utf-8 -*-
"""
Created on Thu Jul 24 11:42:36 2025

@author: 774926
"""

import os
import tempfile
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service


#########

import pyautogui
import shutil
import pandas as pd
import ctypes 
import glob  
# import keyboard
from datetime import datetime
from datetime import timedelta
import re
import time
from selenium.webdriver.common.by import By

prefixo_arquivo = 'Registro de Atendimento'
extensao = ".xls"
numero = 1  # ou qualquer número que você queira usar


dataini = datetime(2025, 11, 1) # Alterar Data Inicial
diaini = dataini.day
mesini = dataini.month
anoini = dataini.year

diaini_formatada = f"{diaini:02d}"
mesini_formatada = f"{mesini:02d}"


datafim = dataini + timedelta(days=90)
diafim = datafim.day
anofim = datafim.year
mesfim = datafim.month
diafim_formatada = f"{diafim:02d}"
mesfim_formatada = f"{mesfim:02d}"


data_ontem = datetime.now() - timedelta(days=1)
print(data_ontem.strftime('%Y-%m-%d'))

print(datafim)


#ontem = datetime.now() - timedelta(days=1)


#user_data_dir = tempfile.mkdtemp()
#PASTA = rf"C:\Users\{os.getlogin()}\Downloads\INTERGRALL" #pasta que irá salvar no seu computador
"""
def configurar_navegador_edge():
    edge_options = Options()

    # Criar diretório temporário para perfil do navegador
    user_data_dir = tempfile.mkdtemp()
    edge_options.add_argument(f"--user-data-dir={user_data_dir}")
    edge_options.add_argument("--start-maximized")

    # CAMINHO MANUAL para o msedgedriver.exe
    driver_path = r"C:\MIS\msedgedriver.exe"  # <-- Use raw string ou escape as barras

    # Criar objeto Service com o caminho do driver
    service = Service(executable_path=driver_path)

    # Inicia o navegador usando o service
    navegador = webdriver.Edge(service=service, options=edge_options)

    return navegador

"""


PASTA = rf"C:\Users\{os.getlogin()}\Downloads\INTERGRALLNAIP" #pasta que irá salvar no seu computador
# Configurando opções do Edge <-- ESTA LINHA APAGA AS CONFIGURAÇÕES ACIMA
edge_options = Options()
edge_options.add_experimental_option("prefs", {
    "download.default_directory": PASTA,
    "download.prompt_for_download": False,
    "directory_upgrade": True,
}
    )

#edge_options = Options()  <-- ESTA LINHA APAGA AS CONFIGURAÇÕES ACIMA
driver_path = r"C:\MIS\msedgedriver.exe"  # <-- Use raw string ou escape as barras
service = Service(executable_path=driver_path)
navegador = webdriver.Edge(service=service, options=edge_options)


# Usar o navegador
#navegador = configurar_navegador_edge()
#navegador.get("https://wwws.intergrall.com.br/callcenter/cc_login.php")

# Percorre todos os arquivos na pasta
def excluir():
    for nome_arquivo in os.listdir(PASTA):
        caminho_completo = os.path.join(PASTA, nome_arquivo)

        # Verifica se é um arquivo
        if os.path.isfile(caminho_completo):
            # Encontra todos os números no nome do arquivo
            numeros = re.findall(r'\d+', nome_arquivo)
            if any(int(num) >= 1 for num in numeros):
                print(f"Removendo: {nome_arquivo}")
                os.remove(caminho_completo)
                
excluir() 

navegador.get("https://wwws.intergrall.com.br/callcenter/cc_login.php")#time.sleep(1)



time.sleep(1)
# -->> DIGITAR LOGIN
# email
navegador.find_element('xpath', '//*[@id="login"]').send_keys("elo1047")
    
    
time.sleep(1)
# -->> DIGITAR SENHA
navegador.find_element('xpath', '//*[@id="password"]').send_keys("baldo2025")  # senha
time.sleep(1)
# -->> DIGITAR ENTRAR
navegador.find_element( 'xpath', '//*[@id="loginBtn"]').click()  # entrar

time.sleep(2)

# -->> DIGITAR SENHA
navegador.find_element('xpath', '//*[@id="password"]').send_keys("baldo2025")  # senha
time.sleep(3)
# -->> DIGITAR ENTRAR
navegador.find_element( 'xpath', '//*[@id="loginBtn"]').click()  # entrar
time.sleep(3)
navegador.find_element( 'xpath', '//*[@id="resBtn"]').click()  # entrar
time.sleep(3)
navegador.find_element( 'xpath', '//*[@id="mainnav"]/ul/li[3]/a').click() 
time.sleep(3)

navegador.find_element( 'xpath', '//*[@id="link_P1I"]').click() 
time.sleep(3)

    
iframe = navegador.find_element(By.ID, "frame_div_aba_1")
navegador.switch_to.frame(iframe)
time.sleep(2)
navegador.find_element( 'xpath', '/html/body/div[2]/form[5]/table/tbody/tr/td[2]').click() 

time.sleep(2)

navegador.find_element( 'xpath', '//*[@id="select2-drop"]/ul/li[2]/div').click() 


navegador.find_element( 'xpath', '/html/body/div[2]/form[5]/table/tbody/tr[2]/td[2]').click() 

pyautogui.press('tab') # Simula seta tb
time.sleep(2)

pyautogui.press('down')  # Simula seta para baixo
time.sleep(2)


navegador.find_element( 'xpath', '//*[@id="exportacao"]').click() 
pyautogui.press('down')  # Simula seta para baixo
time.sleep(2)
pyautogui.press('enter')  # Simula enter




navegador.find_element('xpath', '/html/body/div[2]/form[5]/table/tbody/tr[5]/td[2]/select').click()
pyautogui.press('down')  # Simula seta para baixo
pyautogui.press('down')  # Simula seta para baixo
pyautogui.press('enter')  # Simula enter





time.sleep(2)



while dataini.strftime('%Y-%m-%d') < data_ontem.strftime('%Y-%m-%d'):
    navegador.find_element('xpath', '//*[@id="dia_ini"]').clear()  # senha
    navegador.find_element('xpath', '//*[@id="dia_ini"]').send_keys(diaini_formatada)  # isso força 2 dígitos com zero à esquerda

    time.sleep(1)


    navegador.find_element('xpath', '//*[@id="mes_ini"]').clear()  # senha
    navegador.find_element('xpath', '//*[@id="mes_ini"]').send_keys(mesini_formatada)  # isso força 2 dígitos com zero à esquerda

    time.sleep(1)

    navegador.find_element('xpath', '//*[@id="ano_ini"]').clear()  # senha
    navegador.find_element('xpath', '//*[@id="ano_ini"]').send_keys(anoini)  # isso força 2 dígitos com zero à esquerda

    time.sleep(1)

    navegador.find_element('xpath', '//*[@id="dia_fim"]').send_keys(diafim_formatada)  # isso força 2 dígitos com zero à esquerda
    navegador.find_element('xpath', '//*[@id="dia_fim"]').clear()  # senha


    time.sleep(1)


    navegador.find_element('xpath', '//*[@id="dia_fim"]').clear()  # senha
    navegador.find_element('xpath', '//*[@id="dia_fim"]').send_keys(diafim_formatada)  # isso força 2 dígitos com zero à esquerda

    time.sleep(1)


    navegador.find_element('xpath', '//*[@id="mes_fim"]').clear()  # senha
    navegador.find_element('xpath', '//*[@id="mes_fim"]').send_keys(mesfim_formatada)  # isso força 2 dígitos com zero à esquerda



    time.sleep(1)


    navegador.find_element('xpath', '//*[@id="ano_fim"]').clear()  # senha
    navegador.find_element('xpath', '//*[@id="ano_fim"]').send_keys(anofim ) # isso força 2 dígitos com zero à esquerda

    time.sleep(1)
    navegador.find_element( 'xpath', '/html/body/div[2]/form[5]/div/input').click() 


    time.sleep(1)


    button = navegador.find_element('xpath', '/html/body/div[2]/table/tbody/tr/td/img[2]')
    navegador.execute_script("arguments[0].click();", button)
    time.sleep(3)
    
    dataini = datafim +  timedelta(days=1)
    print(dataini )
    datafim = datafim +  timedelta(days=90)
    if datafim > data_ontem:
        datafim = data_ontem
    timeout = 30
    esperou = 0
    arquivo_encontrado = None

    while esperou < timeout:
        arquivos = os.listdir(PASTA)
        arquivos_filtrados = [f for f in arquivos if f.startswith(prefixo_arquivo) and f.endswith(extensao)]
        
        if arquivos_filtrados:
            # Pega o arquivo mais recente (pode ajustar conforme quiser)
            arquivo_encontrado = sorted(
                arquivos_filtrados, 
                key=lambda f: os.path.getmtime(os.path.join(PASTA, f))
            )[-1]
            break

        time.sleep(1)
        esperou += 1

    if arquivo_encontrado:
        caminho_antigo = os.path.join(PASTA, arquivo_encontrado)
  
        novo_nome = f"Registro de Atendimento ({numero}){extensao}"
        caminho_novo = os.path.join(PASTA, novo_nome)

        os.rename(caminho_antigo, caminho_novo)
        print(f"Arquivo renomeado para: {novo_nome}")
    else:
        print("Arquivo não encontrado dentro do tempo limite.")
            
    
    
    diaini = dataini.day
    mesini = dataini.month
    anoini = dataini.year

    diaini_formatada = f"{diaini:02d}"
    mesini_formatada = f"{mesini:02d}"



    diafim = datafim.day
    anofim = datafim.year
    mesfim = datafim.month
    diafim_formatada = f"{diafim:02d}"
    mesfim_formatada = f"{mesfim:02d}"
    numero = numero + 1
  
    
r"""

Junta todas as bases baixada do Intergrall Registro de Atendimento.xls na pasta:
    .\Downloads\Salesforce
Converte para excel e salva como excel Registro de Atendimento.xlsx na pasta:
    \\brsbesrv960\Publico\REPORTS\001 CARGAS\001 EXCEL\ALELO BKO\OUVIDORIA
"""
time.sleep(4)


print('consolidar')

dir_ = f"{os.environ['USERPROFILE']}\\Downloads\\INTERGRALLNAIP"
dir_out = r'\\brsbesrv960\Publico\REPORTS\001 CARGAS\001 EXCEL\ALELO BKO\NAIP'

colunas = [
    'Usuário Abertura', 'Responsável', 'Protocolo', 'Setor', 'Dt. Abertura', 'Dt. Vecto',
    'Dt. Finalização', 'Origem Atd.', 'Nome', 'CPF/CNPJ', 'Canal', 'Protocolo Canal',
    'Dt. Recebimento', 'CIP/Aud', 'Dt. Aud', 'Manifestação', 'Produto', 'Tipo Cliente',
    'Descrição dos Fatos', 'Valor Envolvido:', 'Valor Dispendido:', 'Prazo Área:',
    'Motivo', 'Submotivo', 'Situação', 'Prorrogado', 'Reiteração', 'Resultado Ouvidoria'
]
colunas_numericas = ['Protocolo', 'Protocolo Canal', 'Valor Envolvido:', 'Valor Dispendido:']
colunas_str = [c for c in colunas if c not in colunas_numericas]
colunas_data = ['Dt. Abertura', 'Dt. Vecto', 'Dt. Finalização', 'Dt. Recebimento', 'Dt. Aud', 'Prazo Área:']

def dateparse(x): 
    return datetime.strptime(x, '%d/%m/%Y')

# Somente .xls (os “HTML disfarçados” do Intergrall)
arquivos_download = [f for f in os.listdir(dir_) if f.lower().endswith('.xls')]
df_final = pd.DataFrame(columns=colunas)

for arq in arquivos_download:
    try:
        # Lê a primeira tabela do arquivo .xls (HTML)
        df = pd.read_html(f"{dir_}\\{arq}", decimal=',', thousands='.', encoding="iso-8859-1")[0]

        # Se o número de colunas não bater com a lista 'colunas', tenta ajustar:
        if df.shape[1] != len(colunas):
            # Tenta usar a primeira linha como header (muito comum nesses HTMLs)
            if df.shape[0] > 0:
                df.columns = df.iloc[0]
                df = df.iloc[1:].reset_index(drop=True)

        # Se ainda assim não bater, vamos selecionar o que der interseção
        if df.shape[1] != len(colunas):
            # mantém apenas as colunas esperadas que existirem
            intersec = [c for c in colunas if c in df.columns]
            df = df[intersec].copy()
            # reindexa para a ordem final, criando as faltantes vazias
            df = df.reindex(columns=colunas)
        else:
            # Se bater bonitinho, força as colunas esperadas
            df.columns = colunas

        # 🔹 AJUSTE OPÇÃO 2: garante dtype string para .str
        df[colunas_str] = df[colunas_str].astype("string")  # <- importante

        # Agora sim, pode usar o .str com encode/decode sem quebrar NaN
        df[colunas_str] = df[colunas_str].apply(
            lambda s: s.str.encode('latin-1', errors='ignore').str.decode('utf-8', errors='replace')
        )

        # 🔹 Datas: substitui 'Não finalizado' por vazio (sem conversão agora)
        df[colunas_data] = df[colunas_data].replace("Não finalizado", "")

        df_final = pd.concat([df_final, df], ignore_index=True)

    except Exception as e:
        print(f"❌ Erro ao processar {arq}: {e}")

# Exporta
out_path = f'{dir_out}\\Registro de Atendimento.xlsx'
df_final.to_excel(out_path, index=False, na_rep='', sheet_name='Registro de Atendimento')
print(f"✅ Consolidado em: {out_path}")
