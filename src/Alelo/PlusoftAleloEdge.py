#type:ignore
# -*- coding: utf-8 -*-
"""
Created on Wed Jul 16 11:34:44 2025

@author: 774926
"""

"""
Created on Thu Jun 26 17:29:21 2025

@author: 774926
"""

## Caminhos utilizados dentro da automação:
# \\brsbesrv960\publico\REPORTS\ALELO\MOTIVO_CHAMADOR_POD - Processo para 1 base
 
# \\brsbesrv960\publico\REPORTS\ALELO\ADQUIRENCIA - Processo para 3 bases
 
# \\brsbesrv960\publico\REPORTS\001 CARGAS\001 EXCEL\ALELO_CHAT_EC - Não mexe
# Procedure: JOB_ALELO_CHAT_EC_MIS31179
 
# \\brsbesrv960\Publico\REPORTS\001 CARGAS\001 EXCEL\ALELO CAU\EMAIL ATENDIMENTO - Processo para 2 bases
"""from cgitb import text"""
from tkinter import messagebox
import requests
import tempfile
from tkinter import *
import  datetime 
from datetime import timedelta
from selenium import webdriver
from  time import sleep
import os
import traceback
import time
import datetime
import re
#from selenium import webdriverp
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import sys

from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import  datetime 
from datetime import timedelta
from time import sleep
"""from webdriver_manager.chrome import ChromeDriverManager #pip install webdriver_manager"""
# import pyautogui
import shutil
import pandas as pd
import ctypes  # An included library with Python install. 
import glob  


#import EdgeOptions

import os
import tempfile
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service


#########

import pyautogui
import shutil
import pandas as pd
import ctypes  # An included library with Python install. 
import glob  
# import keyboard
from datetime import datetime
from datetime import timedelta
import re
import time
from selenium.webdriver.common.by import By




#user_data_dir = tempfile.mkdtemp()
PASTA = rf"C:\Users\{os.getlogin()}\Downloads\CR_DA_CAU" #pasta que irá salvar no seu computador
c = rf"C:\Users\{os.getlogin()}\Downloads\CR_DA_CAU\\"
m = r'\\brsbesrv960\Publico\REPORTS\001 CARGAS\001 EXCEL\ALELO CAU\EMAIL ATENDIMENTO\\' 
m1 = r'\\brsbesrv960\Publico\REPORTS\001 CARGAS\001 EXCEL\ALELO_CHAT_EC\\' 
m2 = r'\\brsbesrv960\publico\REPORTS\ALELO\ADQUIRENCIA\\' 
m3 = r'\\brsbesrv960\Publico\REPORTS\ALELO\MOTIVO_CHAMADOR_POD' 


nome_arquivo_procurado = 'POD CAU - DA - Base E-mail X Atendimento - Concluídos'
nome_arquivo_procurado1 = 'DA - Base E-mail X Atendimento - Novos'
nome_arquivo_procurado2 = 'DA - Base E-mail X Atendimento - Concluídos'
nome_arquivo_procurado3 = 'CR - Base Analítica E-mail X Atendimento - Novos'
nome_arquivo_procurado4 = 'CR - Base Analítica E-mail X Atend. - Concluídos'
nome_arquivo_procurado5 = 'CR - Base Analítica Email Novos'
nome_arquivo_procurado6 = 'CR - Base Analítica Email - Classificados'
nome_arquivo_procurado7 = 'Base Analítica Email Novos'
nome_arquivo_procurado8 = 'Base Analítica Email - Classificados'
nome_arquivo_procurado9 = 'Atendimento Chat EC - Alelo'
nome_arquivo_procurado10 = 'Chat EC - Dados - Operações'
nome_arquivo_procurado11 = 'Pesquisa Chat EC'
nome_arquivo_procurado12 = 'Prevenção Adquirência - Validação de Ajustes'
nome_arquivo_procurado13 = 'Chargeback Adquirência'
nome_arquivo_procurado14 = 'Modelo Preditivo - Fraude'
nome_arquivo_procurado15 = 'Modelo Preditivo - Tickeiro'
nome_arquivo_procurado16 = 'Motivo Chamador - POD'
nome_arquivo_procurado17 = 'Franquia Eliza'
nome_arquivo_procurado18 = 'Email POD CAE - MAY'






lista = os.listdir(c) #lista separando apenas os arquivos do caminho.
print(lista)

#Passo 1 - Define pasta para download
edge_options = Options()
edge_options.add_experimental_option("prefs", {
    "download.default_directory": PASTA,
    "download.prompt_for_download": False,
    "directory_upgrade": True,
})

#edge_options = Options()  <-- ESTA LINHA APAGA AS CONFIGURAÇÕES ACIMA
driver_path = r"C:\MIS\msedgedriver.exe"  # <-- Use raw string ou escape as barras
service = Service(executable_path=driver_path)
navegador = webdriver.Edge(service=service, options=edge_options)




# %%  INICIO DA NAVEGAÇÃO CHOME
# navegador = webdriver.Chrome('C:\MIS\chromedriver.exe')
# navegador = webdriver.Chrome()
#navegador = webdriver.Edge()  # Inicia uma nova instância do navegador Edge
# navegador.maximize_window()
time.sleep(1)


def mover(caminho, destino):
    lista_arquivo = os.listdir(caminho)

    for nome_arquivo in lista_arquivo:
        if nome_arquivo == "teste.txt":  # ou "teste.txt", dependendo do nome exato
            continue  # pula esse arquivo
        origem_completa = os.path.join(caminho, nome_arquivo)
        destino_completo = os.path.join(destino, nome_arquivo)
        shutil.move(origem_completa, destino_completo)
def logar():
# %% ACESSAR O SITE:
#navegador.get("https://alelo.plusoftomni.com.br/#/")
    #navegador.quit() 
    
    navegador.get("https://alelo.plusoftomni.com.br/?m=menu.reportbuilder")#time.sleep(1)
   
    time.sleep(1)
# -->> DIGITAR LOGIN
# email
    navegador.find_element('xpath', '//*[@id="login__username"]').send_keys("isac.brito")
    
    
    time.sleep(1)
# -->> DIGITAR SENHA
    navegador.find_element('xpath', '//*[@id="login__password"]').send_keys("Lilo!@220886")  # senha
    time.sleep(1)
# -->> DIGITAR ENTRAR
    navegador.find_element( 'xpath', '/html/body/div/div[1]/div[1]/div[2]/div/div/div/form/div[2]/div[1]/div[5]/button').click()  # entrar


#pular pro link reportbuilder

#navegador.get("https://alelo.plusoftomni.com.br/?m=menu.reportbuilder")#time.sleep(1)

    time.sleep(5)   

# %%% CLICKAR NA SETA ONDE TEM O NOME
#navegador.find_element(   'xpath', '//*[@id="inpaas-navbar-collapse"]/ul[1]/li/a/small').click()
#time.sleep(1)


def baixarPodCau():

# %%%%  DIGITAR REPORTE BUILDER

###################BAIXAR POD CAU - DA - Base E-mail X Atendimento - Concluídos ##############################
#navegador.find_element('xpath', '//*[@id="inpaas-navbar-collapse"]/ul[1]/li/ul/li[5]/a').click()
#time.sleep(1)
#procurar o frame

    navegador.get("https://alelo.plusoftomni.com.br/?m=menu.reportbuilder")#time.sleep(1)
    time.sleep(3)
    iframe = navegador.find_element(By.ID, "frame_middle")
    navegador.switch_to.frame(iframe)

    time.sleep(2)
#procurar xpath
    navegador.find_element('xpath', '//*[@id="myviews-navbar-collapse"]/ul/li[4]/form/div/input').send_keys("POD CAU - DA - Base E-mail X Atendimento - Concluídos")
    time.sleep(1)
    navegador.find_element('xpath', '//*[@id="myviews-navbar-collapse"]/ul/li[4]/form/div/input').click()
    time.sleep(3)
#clikar na base
    navegador.find_element( 'xpath', '/html/body/div[1]/div/div[2]/div/div[2]/div/a/span[1]').click()  # entrar
    time.sleep(4)
#clikar no botão abrir
    navegador.find_element( 'xpath', '/html/body/div[2]/div[2]/div[3]/button').click()  # entrar
    time.sleep(5)
#navegador.find_element( 'xpath', '//*[@id="div-dropdown-menu"]/button').click()  # entrar
#clikar no botão baixar
    button = navegador.find_element('xpath', '//*[@id="div-dropdown-menu"]/button')
    navegador.execute_script("arguments[0].click();", button)

    time.sleep(3)
    button = navegador.find_element('xpath', '//*[@id="btn-xlsx"]')
    navegador.execute_script("arguments[0].click();", button)
    time.sleep(30)
################### FIM BAIXAR POD CAU - DA - Base E-mail X Atendimento - Concluídos ##############################

def DABaseEmailAtendimentoNovos():

###################BAIXAR BASE DA - Base E-mail X Atendimento - Novos##############################
#pesquisar()
##Clickar minhas visões e pesquisa
    #button = navegador.find_element('xpath', '/html/body/div[1]/div[2]/div[1]/div[2]/button[1]')
    #navegador.execute_script("arguments[0].click();", button)
    navegador.get("https://alelo.plusoftomni.com.br/?m=menu.reportbuilder")#time.sleep(1)
    
    #button = navegador.find_element('xpath', '/html/body/div[1]/div[2]/div[1]/div[2]/button[1]')
    #navegador.execute_script("arguments[0].click();", button)
    time.sleep(2)

#procurar o frame
    iframe = navegador.find_element(By.ID, "frame_middle")
    navegador.switch_to.frame(iframe)


    time.sleep(3)
#procurar xpath
#como já entrou no frame só precisa pesquisar 
    navegador.find_element('xpath', '//*[@id="myviews-navbar-collapse"]/ul/li[4]/form/div/input').send_keys("DA - Base E-mail X Atendimento - Novos")
    time.sleep(1)
    navegador.find_element('xpath', '//*[@id="myviews-navbar-collapse"]/ul/li[4]/form/div/input').click()
##depois criar uma função para apertar em minhas visões e pesquisa
    time.sleep(3)
#clikar na base 
    navegador.find_element( 'xpath', '/html/body/div[1]/div/div[2]/div/div[2]/div/a/span[2]').click()  # entrar
    time.sleep(3)
#clikar na abrir

    navegador.find_element( 'xpath', '/html/body/div[2]/div[2]/div[3]/button').click()  # entrar
    time.sleep(3)

#clikar no botão baixar
    button = navegador.find_element('xpath', '//*[@id="div-dropdown-menu"]/button')
    navegador.execute_script("arguments[0].click();", button)
    time.sleep(3)
#Baixar opção de excel
    button = navegador.find_element('xpath', '//*[@id="btn-xlsx"]')
    navegador.execute_script("arguments[0].click();", button)
    time.sleep(25)

################### FIM BAIXAR BASE DA - Base E-mail X Atendimento - Novos##############################

def DABaseEmailAtendimentoConcluídos():
################### BAIXAR BASE DA - Base E-mail X Atendimento - Concluídos ##############################
#pesquisar()
##Clickar minhas visões e pesquisa
    #button = navegador.find_element('xpath', '/html/body/div[1]/div[2]/div[1]/div[2]/button[1]')
    #navegador.execute_script("arguments[0].click();", button)
    navegador.get("https://alelo.plusoftomni.com.br/?m=menu.reportbuilder")#time.sleep(1)
    
    #button = navegador.find_element('xpath', '/html/body/div[1]/div[2]/div[1]/div[2]/button[1]')
    #navegador.execute_script("arguments[0].click();", button)
    time.sleep(5)

#procurar o frame
    iframe = navegador.find_element(By.ID, "frame_middle")
    navegador.switch_to.frame(iframe)

    time.sleep(5)
#procurar xpath
#como já entrou no frame só precisa pesquisar 
    navegador.find_element('xpath', '//*[@id="myviews-navbar-collapse"]/ul/li[4]/form/div/input').send_keys("DA - Base E-mail X Atendimento - Concluídos")
    time.sleep(1)
    navegador.find_element('xpath', '//*[@id="myviews-navbar-collapse"]/ul/li[4]/form/div/input').click()
##depois criar uma função para apertar em minhas visões e pesquisa
    time.sleep(3)
#clikar na base 
    navegador.find_element( 'xpath', '/html/body/div[1]/div/div[2]/div/div[2]/div/a/span[2]').click()  # entrar
    time.sleep(3)
#clikar na abrir

    navegador.find_element( 'xpath', '/html/body/div[2]/div[2]/div[3]/button').click()  # entrar
    time.sleep(3)

#clikar no botão baixar
    button = navegador.find_element('xpath', '//*[@id="div-dropdown-menu"]/button')
    navegador.execute_script("arguments[0].click();", button)
    time.sleep(5)
#Baixar opção de excel
    button = navegador.find_element('xpath', '//*[@id="btn-xlsx"]')
    navegador.execute_script("arguments[0].click();", button)
    time.sleep(25)
################### FIM BAIXAR DA - Base E-mail X Atendimento - Concluídos ##############################


def CRBaseAnaliticaEmailAtendimentoNovos():
################### CR - Base Analítica E-mail X Atendimento - Novos ##############################
#pesquisar()
##Clickar minhas visões e pesquisa
    #button = navegador.find_element('xpath', '/html/body/div[1]/div[2]/div[1]/div[2]/button[1]')
    #navegador.execute_script("arguments[0].click();", button)
    navegador.get("https://alelo.plusoftomni.com.br/?m=menu.reportbuilder")#time.sleep(1)
    
    #button = navegador.find_element('xpath', '/html/body/div[1]/div[2]/div[1]/div[2]/button[1]')
    #navegador.execute_script("arguments[0].click();", button)
    time.sleep(5)

#procurar o frame
    iframe = navegador.find_element(By.ID, "frame_middle")
    navegador.switch_to.frame(iframe)

#procurar xpath
#como já entrou no frame só precisa pesquisar 
    navegador.find_element('xpath', '//*[@id="myviews-navbar-collapse"]/ul/li[4]/form/div/input').send_keys("CR - Base Analítica E-mail X Atendimento - Novos")
    time.sleep(1)
    navegador.find_element('xpath', '//*[@id="myviews-navbar-collapse"]/ul/li[4]/form/div/input').click()
##depois criar uma função para apertar em minhas visões e pesquisa
    time.sleep(1)
#clikar na base 
    navegador.find_element( 'xpath', '/html/body/div[1]/div/div[2]/div/div[2]/div/a/span[2]').click()  # entrar
    time.sleep(3)
#clikar na abrir

    navegador.find_element( 'xpath', '/html/body/div[2]/div[2]/div[3]/button').click()  # entrar
    time.sleep(3)

#clikar no botão baixar
    button = navegador.find_element('xpath', '//*[@id="div-dropdown-menu"]/button')
    navegador.execute_script("arguments[0].click();", button)
    time.sleep(5)
#Baixar opção de excel
    button = navegador.find_element('xpath', '//*[@id="btn-xlsx"]')
    navegador.execute_script("arguments[0].click();", button)
    time.sleep(15)
################### FIM BAIXAR CR - Base Analítica E-mail X Atendimento - Novos ##############################

def CRBaseAnalíticaEmailAtendConcluidos():
################### CR - Base Analítica E-mail X Atend. - Concluídos ##############################
#pesquisar()
##Clickar minhas visões e pesquisa
    #button = navegador.find_element('xpath', '/html/body/div[1]/div[2]/div[1]/div[2]/button[1]')
    #navegador.execute_script("arguments[0].click();", button)
    navegador.get("https://alelo.plusoftomni.com.br/?m=menu.reportbuilder")#time.sleep(1)
        
        #button = navegador.find_element('xpath', '/html/body/div[1]/div[2]/div[1]/div[2]/button[1]')
        #navegador.execute_script("arguments[0].click();", button)
    time.sleep(5)

    #procurar o frame
    iframe = navegador.find_element(By.ID, "frame_middle")
    navegador.switch_to.frame(iframe)

    time.sleep(5)
#procurar xpath
#como já entrou no frame só precisa pesquisar 
    navegador.find_element('xpath', '//*[@id="myviews-navbar-collapse"]/ul/li[4]/form/div/input').send_keys("CR - Base Analítica E-mail X Atend. - Concluídos")
    time.sleep(1)
    navegador.find_element('xpath', '//*[@id="myviews-navbar-collapse"]/ul/li[4]/form/div/input').click()
##depois criar uma função para apertar em minhas visões e pesquisa
    time.sleep(1)
#clikar na base 
    navegador.find_element( 'xpath', '/html/body/div[1]/div/div[2]/div/div[2]/div/a/span[2]').click()  # entrar
    time.sleep(3)
#clikar na abrir

    navegador.find_element( 'xpath', '/html/body/div[2]/div[2]/div[3]/button').click()  # entrar
    time.sleep(3)

#clikar no botão baixar
    button = navegador.find_element('xpath', '//*[@id="div-dropdown-menu"]/button')
    navegador.execute_script("arguments[0].click();", button)
    time.sleep(5)
#Baixar opção de excel
    button = navegador.find_element('xpath', '//*[@id="btn-xlsx"]')
    navegador.execute_script("arguments[0].click();", button)
    time.sleep(20)

################### FIM BAIXAR CR - Base Analítica E-mail X Atend. - Concluídos ##############################


def CRBaseAnaliticaEmailNovos():
################### CR - Base Analítica Email Novos ##############################
#pesquisar()
##Clickar minhas visões e pesquisa
    #button = navegador.find_element('xpath', '/html/body/div[1]/div[2]/div[1]/div[2]/button[1]')
    #navegador.execute_script("arguments[0].click();", button)
    navegador.get("https://alelo.plusoftomni.com.br/?m=menu.reportbuilder")#time.sleep(1)
    
    #button = navegador.find_element('xpath', '/html/body/div[1]/div[2]/div[1]/div[2]/button[1]')
    #navegador.execute_script("arguments[0].click();", button)
    time.sleep(5)

#procurar o frame
    iframe = navegador.find_element(By.ID, "frame_middle")
    navegador.switch_to.frame(iframe)
    time.sleep(3)
#procurar xpath
#como já entrou no frame só precisa pesquisar 
    navegador.find_element('xpath', '//*[@id="myviews-navbar-collapse"]/ul/li[4]/form/div/input').send_keys("CR - Base Analítica Email Novos")
    time.sleep(1)
    navegador.find_element('xpath', '//*[@id="myviews-navbar-collapse"]/ul/li[4]/form/div/input').click()
##depois criar uma função para apertar em minhas visões e pesquisa
    time.sleep(1)
#clikar na base 
    navegador.find_element( 'xpath', '/html/body/div[1]/div/div[2]/div/div[2]/div/a/span[2]').click()  # entrar
    time.sleep(3)
#clikar na abrir

    navegador.find_element( 'xpath', '/html/body/div[2]/div[2]/div[3]/button').click()  # entrar
    time.sleep(3)

#clikar no botão baixar
    button = navegador.find_element('xpath', '//*[@id="div-dropdown-menu"]/button')
    navegador.execute_script("arguments[0].click();", button)
    time.sleep(3)
#Baixar opção de excel
    button = navegador.find_element('xpath', '//*[@id="btn-xlsx"]')
    navegador.execute_script("arguments[0].click();", button)
    time.sleep(14)
    
################### FIM BAIXAR CR - Base Analítica Email Novos##############################

def CRBaseAnaliticaEmailClassificados():
################### CR - Base Analítica Email - Classificados ##############################
#pesquisar()
##Clickar minhas visões e pesquisa
 #button = navegador.find_element('xpath', '/html/body/div[1]/div[2]/div[1]/div[2]/button[1]')
    #navegador.execute_script("arguments[0].click();", button)
    navegador.get("https://alelo.plusoftomni.com.br/?m=menu.reportbuilder")#time.sleep(1)
    
    #button = navegador.find_element('xpath', '/html/body/div[1]/div[2]/div[1]/div[2]/button[1]')
    #navegador.execute_script("arguments[0].click();", button)
    time.sleep(5)

#procurar o frame
    iframe = navegador.find_element(By.ID, "frame_middle")
    navegador.switch_to.frame(iframe)
    time.sleep(5)
#procurar xpath
#como já entrou no frame só precisa pesquisar 
    navegador.find_element('xpath', '//*[@id="myviews-navbar-collapse"]/ul/li[4]/form/div/input').send_keys("CR - Base Analítica Email - Classificados")
    time.sleep(1)
    navegador.find_element('xpath', '//*[@id="myviews-navbar-collapse"]/ul/li[4]/form/div/input').click()
##depois criar uma função para apertar em minhas visões e pesquisa

    time.sleep(1)
#clikar na base 
    navegador.find_element( 'xpath', '/html/body/div[1]/div/div[2]/div/div[2]/div/a/span[2]').click()  # entrar
    time.sleep(3)
#clikar na abrir

    navegador.find_element( 'xpath', '/html/body/div[2]/div[2]/div[3]/button').click()  # entrar
    time.sleep(5)

#clikar no botão baixar
    button = navegador.find_element('xpath', '//*[@id="div-dropdown-menu"]/button')
    navegador.execute_script("arguments[0].click();", button)
    time.sleep(3)
#Baixar opção de excel
    button = navegador.find_element('xpath', '//*[@id="btn-xlsx"]')
    navegador.execute_script("arguments[0].click();", button)
    time.sleep(15)

################### FIM BAIXAR CR - Base Analítica Email - Classificados##############################

def BaseAnalíticaEmailNovos():
################### Base Analítica Email Novos ##############################
#pesquisar()
##Clickar minhas visões e pesquisa
 #button = navegador.find_element('xpath', '/html/body/div[1]/div[2]/div[1]/div[2]/button[1]')
    #navegador.execute_script("arguments[0].click();", button)
    navegador.get("https://alelo.plusoftomni.com.br/?m=menu.reportbuilder")#time.sleep(1)
    
    #button = navegador.find_element('xpath', '/html/body/div[1]/div[2]/div[1]/div[2]/button[1]')
    #navegador.execute_script("arguments[0].click();", button)
    time.sleep(5)

#procurar o frame
    iframe = navegador.find_element(By.ID, "frame_middle")
    navegador.switch_to.frame(iframe)
#procurar xpath
#como já entrou no frame só precisa pesquisar 
    navegador.find_element('xpath', '//*[@id="myviews-navbar-collapse"]/ul/li[4]/form/div/input').send_keys("Base Analítica Email Novos")
    time.sleep(1)
    navegador.find_element('xpath', '//*[@id="myviews-navbar-collapse"]/ul/li[4]/form/div/input').click()
##depois criar uma função para apertar em minhas visões e pesquisa
    time.sleep(1)
#clikar na base 
    navegador.find_element( 'xpath', '/html/body/div[1]/div/div[2]/div/div[2]/div/a/span[2]').click()  # entrar
    time.sleep(3)
#clikar na abrir

    navegador.find_element( 'xpath', '/html/body/div[2]/div[2]/div[3]/button').click()  # entrar
    time.sleep(5)

#clikar no botão baixar
    button = navegador.find_element('xpath', '//*[@id="div-dropdown-menu"]/button')
    navegador.execute_script("arguments[0].click();", button)
    time.sleep(5)
#Baixar opção de excel
    button = navegador.find_element('xpath', '//*[@id="btn-xlsx"]')
    navegador.execute_script("arguments[0].click();", button)
    time.sleep(15)

################### FIM BAIXAR Base Analítica Email Novos##############################

def BaseAnaliticaEmailClassificados():
################### Base Analítica Email - Classificados ##############################
#pesquisar()
 #button = navegador.find_element('xpath', '/html/body/div[1]/div[2]/div[1]/div[2]/button[1]')
    #navegador.execute_script("arguments[0].click();", button)
    navegador.get("https://alelo.plusoftomni.com.br/?m=menu.reportbuilder")#time.sleep(1)
    
    #button = navegador.find_element('xpath', '/html/body/div[1]/div[2]/div[1]/div[2]/button[1]')
    #navegador.execute_script("arguments[0].click();", button)
    time.sleep(5)

#procurar o frame
    iframe = navegador.find_element(By.ID, "frame_middle")
    navegador.switch_to.frame(iframe)
    time.sleep(5)
#procurar xpath
#como já entrou no frame só precisa pesquisar 
    navegador.find_element('xpath', '//*[@id="myviews-navbar-collapse"]/ul/li[4]/form/div/input').send_keys("Base Analítica Email - Classificados")
    time.sleep(1)
    navegador.find_element('xpath', '//*[@id="myviews-navbar-collapse"]/ul/li[4]/form/div/input').click()
##depois criar uma função para apertar em minhas visões e pesquisa
    time.sleep(1)
#clikar na base 
    navegador.find_element( 'xpath', '/html/body/div[1]/div/div[2]/div/div[2]/div/a/span[2]').click()  # entrar
    time.sleep(3)
#clikar na abrir

    navegador.find_element( 'xpath', '/html/body/div[2]/div[2]/div[3]/button').click()  # entrar
    time.sleep(5)

#clikar no botão baixar
    button = navegador.find_element('xpath', '//*[@id="div-dropdown-menu"]/button')
    navegador.execute_script("arguments[0].click();", button)
    time.sleep(5)
#Baixar opção de excel
    button = navegador.find_element('xpath', '//*[@id="btn-xlsx"]')
    navegador.execute_script("arguments[0].click();", button)

    time.sleep(17)
################### FIM BAIXAR Base Analítica Email - Classificados##############################


################### FIM BAIXAR Base Analítica Email Novos##############################

def AtendimentoChatECAlelo():
###################Atendimento Chat EC - Alelo##############################
#pesquisar()
##Clickar minhas visões e pesquisa
    #button = navegador.find_element('xpath', '/html/body/div[1]/div[2]/div[1]/div[2]/button[1]')
    #navegador.execute_script("arguments[0].click();", button)
    #time.sleep(5)
#procurar xpath
#como já entrou no frame só precisa pesquisar 
    navegador.get("https://alelo.plusoftomni.com.br/?m=menu.reportbuilder")#time.sleep(1)
    
    #button = navegador.find_element('xpath', '/html/body/div[1]/div[2]/div[1]/div[2]/button[1]')
    #navegador.execute_script("arguments[0].click();", button)
    time.sleep(5)

#procurar o frame
    iframe = navegador.find_element(By.ID, "frame_middle")
    navegador.switch_to.frame(iframe)
    
    navegador.find_element('xpath', '//*[@id="myviews-navbar-collapse"]/ul/li[4]/form/div/input').send_keys("Atendimento Chat EC - Alelo")
    time.sleep(1)
    navegador.find_element('xpath', '//*[@id="myviews-navbar-collapse"]/ul/li[4]/form/div/input').click()
##depois criar uma função para apertar em minhas visões e pesquisa
    time.sleep(1)
#clikar na base 
    navegador.find_element( 'xpath', '/html/body/div[1]/div/div[2]/div/div[2]/div/a/span[2]').click()  # entrar
    time.sleep(3)
#clikar na abrir

    navegador.find_element( 'xpath', '/html/body/div[2]/div[2]/div[3]/button').click()  # entrar
    time.sleep(5)

#clikar no botão baixar
    button = navegador.find_element('xpath', '//*[@id="div-dropdown-menu"]/button')
    navegador.execute_script("arguments[0].click();", button)
    time.sleep(5)
#Baixar opção de excel
    button = navegador.find_element('xpath', '//*[@id="btn-xlsx"]')
    navegador.execute_script("arguments[0].click();", button)

    time.sleep(17)
################### FIM BAIXAR Base Atendimento Chat EC - Alelo##############################


def ChatECDadosOperações():
###################Atendimento Chat EC - Alelo##############################
#pesquisar()
##Clickar minhas visões e pesquisa
#    button = navegador.find_element('xpath', '/html/body/div[1]/div[2]/div[1]/div[2]/button[1]')
#   navegador.execute_script("arguments[0].click();", button)
    navegador.get("https://alelo.plusoftomni.com.br/?m=menu.reportbuilder")#time.sleep(1)
    
    #button = navegador.find_element('xpath', '/html/body/div[1]/div[2]/div[1]/div[2]/button[1]')
    #navegador.execute_script("arguments[0].click();", button)
    time.sleep(5)

#procurar o frame
    iframe = navegador.find_element(By.ID, "frame_middle")
    navegador.switch_to.frame(iframe)
    time.sleep(5)
#procurar xpath
#como já entrou no frame só precisa pesquisar 
    navegador.find_element('xpath', '//*[@id="myviews-navbar-collapse"]/ul/li[4]/form/div/input').send_keys("Chat EC - Dados - Operações")
    time.sleep(1)
    navegador.find_element('xpath', '//*[@id="myviews-navbar-collapse"]/ul/li[4]/form/div/input').click()
##depois criar uma função para apertar em minhas visões e pesquisa
    time.sleep(1)
#clikar na base 
    navegador.find_element( 'xpath', '/html/body/div[1]/div/div[2]/div/div[2]/div/a/span[2]').click()  # entrar
    time.sleep(3)
#clikar na abrir

    navegador.find_element( 'xpath', '/html/body/div[2]/div[2]/div[3]/button').click()  # entrar
    time.sleep(5)

#clikar no botão baixar
    button = navegador.find_element('xpath', '//*[@id="div-dropdown-menu"]/button')
    navegador.execute_script("arguments[0].click();", button)
    time.sleep(3)
#Baixar opção de excel
    button = navegador.find_element('xpath', '//*[@id="btn-xlsx"]')
    navegador.execute_script("arguments[0].click();", button)

    time.sleep(17)
################### FIM BAIXAR Base Atendimento Chat EC - Alelo##############################



def PesquisaChatEC():
###################Atendimento Chat EC - Alelo##############################
#pesquisar()
##Clickar minhas visões e pesquisa
    navegador.get("https://alelo.plusoftomni.com.br/?m=menu.reportbuilder")#time.sleep(1)
    
    #button = navegador.find_element('xpath', '/html/body/div[1]/div[2]/div[1]/div[2]/button[1]')
    #navegador.execute_script("arguments[0].click();", button)
    time.sleep(5)

#procurar o frame
    iframe = navegador.find_element(By.ID, "frame_middle")
    navegador.switch_to.frame(iframe)

#procurar xpath
#como já entrou no frame só precisa pesquisar 
    navegador.find_element('xpath', '//*[@id="myviews-navbar-collapse"]/ul/li[4]/form/div/input').send_keys("Pesquisa Chat EC")
    time.sleep(1)
    navegador.find_element('xpath', '//*[@id="myviews-navbar-collapse"]/ul/li[4]/form/div/input').click()
##depois criar uma função para apertar em minhas visões e pesquisa
    time.sleep(3)
#clikar na base 
    navegador.find_element( 'xpath', '/html/body/div[1]/div/div[2]/div/div[2]/div/a/span[2]').click()  # entrar
    time.sleep(3)
#clikar na abrir

    navegador.find_element( 'xpath', '/html/body/div[2]/div[2]/div[3]/button').click()  # entrar
    time.sleep(8)

#clikar no botão baixar
    button = navegador.find_element('xpath', '//*[@id="div-dropdown-menu"]/button')
    navegador.execute_script("arguments[0].click();", button)
    time.sleep(8)
#Baixar opção de excel
    button = navegador.find_element('xpath', '//*[@id="btn-xlsx"]')
    navegador.execute_script("arguments[0].click();", button)

    time.sleep(25)
################### FIM BAIXAR Base Atendimento Chat EC - Alelo##############################




def PrevençãoAdquirenciaValidaçãoAjustes():
################### Prevenção Adquirência - Validação de Ajustes ##############################
#pesquisar()
##Clickar minhas visões e pesquisa
    navegador.get("https://alelo.plusoftomni.com.br/?m=menu.reportbuilder")#time.sleep(1)
    
    #button = navegador.find_element('xpath', '/html/body/div[1]/div[2]/div[1]/div[2]/button[1]')
    #navegador.execute_script("arguments[0].click();", button)
    time.sleep(5)

#procurar o frame
    iframe = navegador.find_element(By.ID, "frame_middle")
    navegador.switch_to.frame(iframe)

#procurar xpath
#como já entrou no frame só precisa pesquisar 
    navegador.find_element('xpath', '//*[@id="myviews-navbar-collapse"]/ul/li[4]/form/div/input').send_keys("Prevenção Adquirência - Validação de Ajustes")
    time.sleep(1)
    navegador.find_element('xpath', '//*[@id="myviews-navbar-collapse"]/ul/li[4]/form/div/input').click()
##depois criar uma função para apertar em minhas visões e pesquisa
    time.sleep(1)
#clikar na base 
    navegador.find_element( 'xpath', '/html/body/div[1]/div/div[2]/div/div[2]/div/a/span[2]').click()  # entrar
    time.sleep(3)
#clikar na abrir

    navegador.find_element( 'xpath', '/html/body/div[2]/div[2]/div[3]/button').click()  # entrar
    time.sleep(6)

#clikar no botão baixar
    button = navegador.find_element('xpath', '//*[@id="div-dropdown-menu"]/button')
    navegador.execute_script("arguments[0].click();", button)
    time.sleep(5)
#Baixar opção de excel
    button = navegador.find_element('xpath', '//*[@id="btn-xlsx"]')
    navegador.execute_script("arguments[0].click();", button)

    time.sleep(15)
################### FIM BAIXAR Prevenção Adquirência - Validação de Ajustes ##############################


def ChargebackAdquirencia():
################### Prevenção Adquirência - Validação de Ajustes ##############################
#pesquisar()
##Clickar minhas visões e pesquisa
    navegador.get("https://alelo.plusoftomni.com.br/?m=menu.reportbuilder")#time.sleep(1)
    
    #button = navegador.find_element('xpath', '/html/body/div[1]/div[2]/div[1]/div[2]/button[1]')
    #navegador.execute_script("arguments[0].click();", button)
    time.sleep(5)

#procurar o frame
    iframe = navegador.find_element(By.ID, "frame_middle")
    navegador.switch_to.frame(iframe)

#procurar xpath
#como já entrou no frame só precisa pesquisar 
    navegador.find_element('xpath', '//*[@id="myviews-navbar-collapse"]/ul/li[4]/form/div/input').send_keys("Chargeback Adquirência")
    time.sleep(1)
    navegador.find_element('xpath', '//*[@id="myviews-navbar-collapse"]/ul/li[4]/form/div/input').click()
##depois criar uma função para apertar em minhas visões e pesquisa
    time.sleep(1)
#clikar na base 
    navegador.find_element( 'xpath', '/html/body/div[1]/div/div[2]/div/div[2]/div/a/span[2]').click()  # entrar
    time.sleep(3)
#clikar na abrir

    navegador.find_element( 'xpath', '/html/body/div[2]/div[2]/div[3]/button').click()  # entrar
    time.sleep(3)

#clikar no botão baixar
    button = navegador.find_element('xpath', '//*[@id="div-dropdown-menu"]/button')
    navegador.execute_script("arguments[0].click();", button)
    time.sleep(3)
#Baixar opção de excel
    button = navegador.find_element('xpath', '//*[@id="btn-xlsx"]')
    navegador.execute_script("arguments[0].click();", button)

    time.sleep(10)
################### FIM BAIXAR Prevenção Adquirência - Validação de Ajustes ##############################

def ModeloPreditivoFraude():
################### Prevenção Adquirência - Validação de Ajustes ##############################
#pesquisar()
##Clickar minhas visões e pesquisa
    navegador.get("https://alelo.plusoftomni.com.br/?m=menu.reportbuilder")#time.sleep(1)
    
    #button = navegador.find_element('xpath', '/html/body/div[1]/div[2]/div[1]/div[2]/button[1]')
    #navegador.execute_script("arguments[0].click();", button)
    time.sleep(5)

#procurar o frame
    iframe = navegador.find_element(By.ID, "frame_middle")
    navegador.switch_to.frame(iframe)

#procurar xpath
#como já entrou no frame só precisa pesquisar 
    navegador.find_element('xpath', '//*[@id="myviews-navbar-collapse"]/ul/li[4]/form/div/input').send_keys("Modelo Preditivo - Fraude")
    time.sleep(1)
    navegador.find_element('xpath', '//*[@id="myviews-navbar-collapse"]/ul/li[4]/form/div/input').click()
##depois criar uma função para apertar em minhas visões e pesquisa
    time.sleep(1)
#clikar na base 
    navegador.find_element( 'xpath', '/html/body/div[1]/div/div[2]/div/div[2]/div/a/span[2]').click()  # entrar
    time.sleep(3)
#clikar na abrir

    navegador.find_element( 'xpath', '/html/body/div[2]/div[2]/div[3]/button').click()  # entrar
    time.sleep(7)

#clikar no botão baixar
    button = navegador.find_element('xpath', '//*[@id="div-dropdown-menu"]/button')
    navegador.execute_script("arguments[0].click();", button)
    time.sleep(3)
#Baixar opção de excel
    button = navegador.find_element('xpath', '//*[@id="btn-xlsx"]')
    navegador.execute_script("arguments[0].click();", button)

    time.sleep(10)
################### FIM BAIXAR Modelo Preditivo - Fraude ##############################



def ModeloPreditivoTickeiro():
################### Prevenção Adquirência - Validação de Ajustes ##############################
#pesquisar()
##Clickar minhas visões e pesquisa
    navegador.get("https://alelo.plusoftomni.com.br/?m=menu.reportbuilder")#time.sleep(1)
    
    #button = navegador.find_element('xpath', '/html/body/div[1]/div[2]/div[1]/div[2]/button[1]')
    #navegador.execute_script("arguments[0].click();", button)
    time.sleep(2)

#procurar o frame
    iframe = navegador.find_element(By.ID, "frame_middle")
    navegador.switch_to.frame(iframe)

#procurar xpath
#como já entrou no frame só precisa pesquisar 
    navegador.find_element('xpath', '//*[@id="myviews-navbar-collapse"]/ul/li[4]/form/div/input').send_keys("Modelo Preditivo - Tickeiro")
    time.sleep(1)
    navegador.find_element('xpath', '//*[@id="myviews-navbar-collapse"]/ul/li[4]/form/div/input').click()

##depois criar uma função para apertar em minhas visões e pesquisa
    time.sleep(1)
#clikar na base 
    navegador.find_element( 'xpath', '/html/body/div[1]/div/div[2]/div/div[2]/div/a/span[2]').click()  # entrar
    time.sleep(3)
#clikar na abrir

    navegador.find_element( 'xpath', '/html/body/div[2]/div[2]/div[3]/button').click()  # entrar
    time.sleep(7)

#clikar no botão baixar
    button = navegador.find_element('xpath', '//*[@id="div-dropdown-menu"]/button')
    navegador.execute_script("arguments[0].click();", button)
    time.sleep(3)
#Baixar opção de excel
    button = navegador.find_element('xpath', '//*[@id="btn-xlsx"]')
    navegador.execute_script("arguments[0].click();", button)

    time.sleep(10)
################### FIM BAIXAR Modelo Preditivo - Fraude ##############################



def MotivoChamadorPOD():
################### Prevenção Adquirência - Validação de Ajustes ##############################
#pesquisar()
##Clickar minhas visões e pesquisa
    navegador.get("https://alelo.plusoftomni.com.br/?m=menu.reportbuilder")#time.sleep(1)
    
    #button = navegador.find_element('xpath', '/html/body/div[1]/div[2]/div[1]/div[2]/button[1]')
    #navegador.execute_script("arguments[0].click();", button)
    time.sleep(2)

#procurar o frame
    iframe = navegador.find_element(By.ID, "frame_middle")
    navegador.switch_to.frame(iframe)

#procurar xpath
#como já entrou no frame só precisa pesquisar 
    navegador.find_element('xpath', '//*[@id="myviews-navbar-collapse"]/ul/li[4]/form/div/input').send_keys("Motivo Chamador - POD")
    time.sleep(1)
    navegador.find_element('xpath', '//*[@id="myviews-navbar-collapse"]/ul/li[4]/form/div/input').click()
##depois criar uma função para apertar em minhas visões e pesquisa
    time.sleep(1)
#clikar na base 
    navegador.find_element( 'xpath', '/html/body/div[1]/div/div[2]/div/div[2]/div/a/span[2]').click()  # entrar
    time.sleep(3)
#clikar na abrir

    navegador.find_element( 'xpath', '/html/body/div[2]/div[2]/div[3]/button').click()  # entrar
    time.sleep(3)

#clikar no botão baixar
    button = navegador.find_element('xpath', '//*[@id="div-dropdown-menu"]/button')
    navegador.execute_script("arguments[0].click();", button)
    time.sleep(3)
#Baixar opção de excel
    button = navegador.find_element('xpath', '//*[@id="btn-xlsx"]')
    navegador.execute_script("arguments[0].click();", button)

    time.sleep(25)
################### FIM BAIXAR Motivo Chamador POD ##############################

def FranquiaEliza():
################### Franquia Eliza ##############################
#pesquisar()
##Clickar minhas visões e pesquisa
    navegador.get("https://alelo.plusoftomni.com.br/?m=menu.reportbuilder")#time.sleep(1)
    
    #button = navegador.find_element('xpath', '/html/body/div[1]/div[2]/div[1]/div[2]/button[1]')
    #navegador.execute_script("arguments[0].click();", button)
    time.sleep(2)

#procurar o frame
    iframe = navegador.find_element(By.ID, "frame_middle")
    navegador.switch_to.frame(iframe)

#procurar xpath
#como já entrou no frame só precisa pesquisar 
    navegador.find_element('xpath', '//*[@id="myviews-navbar-collapse"]/ul/li[4]/form/div/input').send_keys("Franquia Eliza")
    time.sleep(1)
    navegador.find_element('xpath', '//*[@id="myviews-navbar-collapse"]/ul/li[4]/form/div/input').click()
##depois criar uma função para apertar em minhas visões e pesquisa
    time.sleep(1)
#clikar na base 
    navegador.find_element( 'xpath', '/html/body/div[1]/div/div[2]/div/div[2]/div/a/span[2]').click()  # entrar
    time.sleep(3)
#clikar na abrir

    navegador.find_element( 'xpath', '/html/body/div[2]/div[2]/div[3]/button').click()  # entrar
    time.sleep(3)

#clikar no botão baixar
    button = navegador.find_element('xpath', '//*[@id="div-dropdown-menu"]/button')
    navegador.execute_script("arguments[0].click();", button)
    time.sleep(3)
#Baixar opção de excel
    button = navegador.find_element('xpath', '//*[@id="btn-xlsx"]')
    navegador.execute_script("arguments[0].click();", button)

    time.sleep(25)
################### FIM BAIXAR Franquia Eliza ##############################

def EmailCaePodMay():
################### Email POD CAE - MAY ##############################
#pesquisar()
##Clickar minhas visões e pesquisa
    navegador.get("https://alelo.plusoftomni.com.br/?m=menu.reportbuilder")#time.sleep(1)
    
    #button = navegador.find_element('xpath', '/html/body/div[1]/div[2]/div[1]/div[2]/button[1]')
    #navegador.execute_script("arguments[0].click();", button)
    time.sleep(2)

#procurar o frame
    iframe = navegador.find_element(By.ID, "frame_middle")
    navegador.switch_to.frame(iframe)

#procurar xpath
#como já entrou no frame só precisa pesquisar 
    navegador.find_element('xpath', '//*[@id="myviews-navbar-collapse"]/ul/li[4]/form/div/input').send_keys("Email POD CAE - MAY")
    time.sleep(1)
    navegador.find_element('xpath', '//*[@id="myviews-navbar-collapse"]/ul/li[4]/form/div/input').click()
##depois criar uma função para apertar em minhas visões e pesquisa
    time.sleep(1)
#clikar na base 
    navegador.find_element( 'xpath', '/html/body/div[1]/div/div[2]/div/div[2]/div/a/span[2]').click()  # entrar
    time.sleep(3)
#clikar na abrir

    navegador.find_element( 'xpath', '/html/body/div[2]/div[2]/div[3]/button').click()  # entrar
    time.sleep(3)

#clikar no botão baixar
    button = navegador.find_element('xpath', '//*[@id="div-dropdown-menu"]/button')
    navegador.execute_script("arguments[0].click();", button)
    time.sleep(3)
#Baixar opção de excel
    button = navegador.find_element('xpath', '//*[@id="btn-xlsx"]')
    navegador.execute_script("arguments[0].click();", button)

    time.sleep(25)
################### FIM BAIXAR Email POD CAE - MAY ##############################

# Variável para armazenar o caminho do arquivo encontrado
caminho_arquivo_encontrado = None
localizadoarquivo = []

# Procurar o arquivo no diretório de origem


# Cria um arquivo txt vazio



#print(PASTA + '\\' + 'teste.txt')

def criararquivo():
    with open(PASTA + '\\' + 'teste.txt', 'w') as arquivo:
        pass  # Não adiciona nada ao arquivo, apenas o cria

criararquivo()
#print(localizadoarquivo)
logar()
#print(max(os.listdir(c)))
def procurararquivo(nome_arquivo_procurado): 
    for arquivo in os.listdir(c):    
        if  arquivo.startswith(nome_arquivo_procurado) :
            print('localizado')
            #break
            localizadoarquivo.append('localizado')
        else:
        #print('teste')
        #logar()
        #baixarPodCau()
        #print('começou abaixar')    
        #time.sleep('15')
        #arquivo = ''
             localizadoarquivo.append('Não localizado')

def procurararquivorede(nome_arquivo_procurado,rede): 
    for arquivo in os.listdir(m):    
        if  arquivo.startswith(nome_arquivo_procurado) :
            print('localizado')
            #break
            localizadoarquivo.append('localizado')
        else:
        #print('teste')
        #logar()
        #baixarPodCau()
        #print('começou abaixar')    
        #time.sleep('15')
        #arquivo = ''
             localizadoarquivo.append('Não localizado')
    

#Certifique-se de substituir 

#procurararquivo(nome_arquivo_procurado)

#print(max(localizadoarquivo))
def baixarPodCauDef():
    while max(localizadoarquivo) != 'localizado' :
            #logar()
            baixarPodCau()
            print('baixou teste Download')
            time.sleep(5)
            procurararquivo(nome_arquivo_procurado) #chamar a lista para verificar os arquivoas atualizados  
            
            print('proxima base') 

#localizadoarquivo = []   
#procurararquivo(nome_arquivo_procurado1)
def DABaseEmailAtendimentoNovosdef():
    while max(localizadoarquivo) != 'localizado' :
            #logar()
            DABaseEmailAtendimentoNovos()
            print('baixou')
            procurararquivo(nome_arquivo_procurado1) #chamar a lista para verificar os arquivoas atualizados  



#print(localizadoarquivo)
def DABaseEmailAtendimentoConcluídosdef():
    while max(localizadoarquivo) != 'localizado' :
            #print('caiu no loop'+localizadoarquivo)
            #logar()
            DABaseEmailAtendimentoConcluídos()
            #print('Não caiu no loop'+localizadoarquivo)
            
            procurararquivo(nome_arquivo_procurado2) #chamar a lista para verificar os arquivoas atualizados  


#print(localizadoarquivo)
def CRBaseAnaliticaEmailAtendimentoNovosdef():
    while max(localizadoarquivo) != 'localizado' :
            #print('caiu no loop'+localizadoarquivo)
            #logar()
            CRBaseAnaliticaEmailAtendimentoNovos()
            #print('Não caiu no loop'+localizadoarquivo)
            
            procurararquivo(nome_arquivo_procurado3) #chamar a lista para verificar os arquivoas atualizados  

def CRBaseAnalíticaEmailAtendConcluidosdef():
    while max(localizadoarquivo) != 'localizado' :
            #print('caiu no loop'+localizadoarquivo)
            #logar()
            CRBaseAnalíticaEmailAtendConcluidos()
            #print('Não caiu no loop'+localizadoarquivo)
            
            procurararquivo(nome_arquivo_procurado4) #chamar a lista para verificar os arquivoas atualizados  

def CRBaseAnaliticaEmailNovosdf():
    while max(localizadoarquivo) != 'localizado' :
            #print('caiu no loop'+localizadoarquivo)
            #logar()
            CRBaseAnaliticaEmailNovos()
            #print('Não caiu no loop'+localizadoarquivo)
            
            procurararquivo(nome_arquivo_procurado5) #chamar a lista para verificar os arquivoas atualizados  

def CRBaseAnaliticaEmailClassificadosdf():
    while max(localizadoarquivo) != 'localizado' :
            #print('caiu no loop'+localizadoarquivo)
            #logar()
            CRBaseAnaliticaEmailClassificados()
            #print('Não caiu no loop'+localizadoarquivo)
            
            procurararquivo(nome_arquivo_procurado6) #chamar a lista para verificar os arquivoas atualizados  

def BaseAnalíticaEmailNovosdf():
    while max(localizadoarquivo) != 'localizado' :
            #print('caiu no loop'+localizadoarquivo)
            #logar()
            BaseAnalíticaEmailNovos()
            #print('Não caiu no loop'+localizadoarquivo)
            
            procurararquivo(nome_arquivo_procurado7) #chamar a lista para verificar os arquivoas atualizados  

def BaseAnaliticaEmailClassificadosdf():
    while max(localizadoarquivo) != 'localizado' :
            #print('caiu no loop'+localizadoarquivo)
            #logar()
            BaseAnaliticaEmailClassificados()
            #print('Não caiu no loop'+localizadoarquivo)
            
            procurararquivo(nome_arquivo_procurado8) #chamar a lista para verificar os arquivoas atualizados  
def AtendimentoChatECAlelodf():
    while max(localizadoarquivo) != 'localizado' :
            #print('caiu no loop'+localizadoarquivo)
            #logar()
            AtendimentoChatECAlelo()
            #print('Não caiu no loop'+localizadoarquivo)
            
            procurararquivo(nome_arquivo_procurado9) #chamar a lista para verificar os arquivoas atualizados  

def ChatECDadosOperaçõesdf():
    while max(localizadoarquivo) != 'localizado' :
            #print('caiu no loop'+localizadoarquivo)
            #logar()
            ChatECDadosOperações()
            #print('Não caiu no loop'+localizadoarquivo)
            
            procurararquivo(nome_arquivo_procurado10) #chamar a lista para verificar os arquivoas atualizados  

def PesquisaChatECdf():
    while max(localizadoarquivo) != 'localizado' :
            #print('caiu no loop'+localizadoarquivo)
            #logar()
            PesquisaChatEC()
            #print('Não caiu no loop'+localizadoarquivo)
            
            procurararquivo(nome_arquivo_procurado11) #chamar a lista para verificar os arquivoas atualizados  

def PrevençãoAdquirenciaValidaçãoAjustesdf():
    while max(localizadoarquivo) != 'localizado' :
            #print('caiu no loop'+localizadoarquivo)
            #logar()
            PrevençãoAdquirenciaValidaçãoAjustes()
            #print('Não caiu no loop'+localizadoarquivo)
            
            procurararquivo(nome_arquivo_procurado12) #chamar a lista para verificar os arquivoas atualizados  

def ChargebackAdquirenciadf():
    while max(localizadoarquivo) != 'localizado' :
            #print('caiu no loop'+localizadoarquivo)
            #logar()
            ChargebackAdquirencia()
            #print('Não caiu no loop'+localizadoarquivo)
            
            procurararquivo(nome_arquivo_procurado13) #chamar a lista para verificar os arquivoas atualizados  

def ModeloPreditivoFraudedf():
    while max(localizadoarquivo) != 'localizado' :
            #print('caiu no loop'+localizadoarquivo)
            #logar()
            ModeloPreditivoFraude()
            #print('Não caiu no loop'+localizadoarquivo)
            
            procurararquivo(nome_arquivo_procurado14) #chamar a lista para verificar os arquivoas atualizados  


def ModeloPreditivoTickeirodf():
    while max(localizadoarquivo) != 'localizado' :
            #print('caiu no loop'+localizadoarquivo)
            #logar()
            ModeloPreditivoTickeiro()
            #print('Não caiu no loop'+localizadoarquivo)
            
            procurararquivo(nome_arquivo_procurado15) #chamar a lista para verificar os arquivoas atualizados  


def MotivoChamadorPOdf():
    while max(localizadoarquivo) != 'localizado' :
            #print('caiu no loop'+localizadoarquivo)
            #logar()
            MotivoChamadorPOD()
            #print('Não caiu no loop'+localizadoarquivo)
            
            procurararquivo(nome_arquivo_procurado16) #chamar a lista para verificar os arquivoas atualizados  

def FranquiaElizadf():
    while max(localizadoarquivo) != 'localizado' :
            #print('caiu no loop'+localizadoarquivo)
            #logar()
            FranquiaEliza()
            #print('Não caiu no loop'+localizadoarquivo)
            
            procurararquivo(nome_arquivo_procurado17) #chamar a lista para verificar os arquivoas atualizados  
            
def EmailCaePodMaydf():
    while max(localizadoarquivo) != 'localizado' :
            #print('caiu no loop'+localizadoarquivo)
            #logar()
            EmailCaePodMay()
            #print('Não caiu no loop'+localizadoarquivo)
            
            procurararquivo(nome_arquivo_procurado18) #chamar a lista para verificar os arquivoas atualizados  

#localizadoarquivo = []   
#procurararquivo(nome_arquivo_procurado1)

#print(max(localizadoarquivo))
#print(nome_arquivo_procurado3)

##INICIO \\brsbesrv960\Publico\REPORTS\001 CARGAS\001 EXCEL\ALELO CAU\EMAIL ATENDIMENTO          
####################################################################################################
########################################################################################################
#localizadoarquivo = [] 


procurararquivo(nome_arquivo_procurado)

baixarPodCauDef()#baixar 
localizadoarquivo = []   
procurararquivo(nome_arquivo_procurado1)
#mover(c, m)
#criararquivo()
DABaseEmailAtendimentoNovosdef()#baixar 
#time.sleep(10)
localizadoarquivo = []   
procurararquivo(nome_arquivo_procurado2)


DABaseEmailAtendimentoConcluídosdef() #baixar 
localizadoarquivo = []   
procurararquivo(nome_arquivo_procurado3)

CRBaseAnaliticaEmailAtendimentoNovosdef()
localizadoarquivo = []   
procurararquivo(nome_arquivo_procurado4)

CRBaseAnalíticaEmailAtendConcluidosdef()
localizadoarquivo = []   
procurararquivo(nome_arquivo_procurado5)

CRBaseAnaliticaEmailNovosdf()
localizadoarquivo = []   
procurararquivo(nome_arquivo_procurado6)


CRBaseAnaliticaEmailClassificadosdf()
localizadoarquivo = []   
procurararquivo(nome_arquivo_procurado7)


BaseAnalíticaEmailNovosdf()
localizadoarquivo = []   
procurararquivo(nome_arquivo_procurado8)


BaseAnaliticaEmailClassificadosdf()
localizadoarquivo = []
procurararquivo(nome_arquivo_procurado17)


FranquiaElizadf()
localizadoarquivo = []
procurararquivo(nome_arquivo_procurado18)

EmailCaePodMaydf()
if len(localizadoarquivo) >= 12:
    #print('Mover')
    mover(c, m)
    #print('Mover')
else:
    print('Falta arquivo Rede EMAIL ATENDIMENTO')
    sys.exit()


##Fim \\brsbesrv960\Publico\REPORTS\001 CARGAS\001 EXCEL\ALELO CAU\EMAIL ATENDIMENTO          
##############################################################################################
########################################################################################


## INICIO \\brsbesrv960\Publico\REPORTS\001 CARGAS\001 EXCEL\ALELO_CHAT_EC  
##############################################################################################
########################################################################################


localizadoarquivo = []   
procurararquivo(nome_arquivo_procurado9)
AtendimentoChatECAlelodf()

localizadoarquivo = []   
procurararquivo(nome_arquivo_procurado10)        
    
ChatECDadosOperaçõesdf()
localizadoarquivo = []   
procurararquivo(nome_arquivo_procurado11)


PesquisaChatECdf()


if len(localizadoarquivo) >= 4:
    #print('Mover')
    mover(c, m1)
    #print('Mover')
else:
    print('Falta arquivo Rede EMAIL ATENDIMENTO')
    sys.exit()


######### FIM  \\brsbesrv960\Publico\REPORTS\001 CARGAS\001 EXCEL\ALELO_CHAT_EC    

###########################\\brsbesrv960\publico\REPORTS\ALELO\ADQUIRENCIA##########################
##print(len(localizadoarquivo))
localizadoarquivo = []   
procurararquivo(nome_arquivo_procurado12)
PrevençãoAdquirenciaValidaçãoAjustesdf()
localizadoarquivo = []   
procurararquivo(nome_arquivo_procurado13)


ChargebackAdquirenciadf()
localizadoarquivo = []   
procurararquivo(nome_arquivo_procurado14)



ModeloPreditivoFraudedf()
localizadoarquivo = []   
procurararquivo(nome_arquivo_procurado15)


ModeloPreditivoTickeirodf()

if len(localizadoarquivo) >= 5:
    #print('Mover')
    mover(c, m2)
    #print('Mover')
else:
    print('Falta arquivo Rede ADQUIRENCIA')
    sys.exit()
    
    
 #########################################################################   
    
   
    
################################# MotivoChamadorPO  #######################
###########################################################################  
    
localizadoarquivo = []   
procurararquivo(nome_arquivo_procurado16)
MotivoChamadorPOdf()   

if len(localizadoarquivo) >= 1:
    #print('Mover')
    mover(c, m3)
    #print('Mover')
else:
    print('Falta arquivo Rede ADQUIRENCIA')
    sys.exit()

  ##############################################################################  