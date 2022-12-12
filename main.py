import pyautogui as gui
import webbrowser as wbrowser
import time 
from selenium import webdriver as wdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd

# preencher
urlLogin = 'http://helpdesk.com/login.php'
usuario = ''
senha = ''
urlAbrirChamado = 'http://helpdesk.com/ocomon/geral/ticket_add.php'

gui.alert("vai começar... ")



#abrir navegador

browser= wdriver.Firefox()
# url de login 
browser.get(urlLogin)
time.sleep(3)
browser.find_element(By.ID,'user').send_keys(usuario)
browser.find_element(By.ID,'pass').send_keys(senha)
time.sleep(3)
browser.find_element(By.ID,'bt_login').click()
time.sleep(3)

#obter dados do arquivo excel
dataframe = pd.read_excel("./ocorrencias.xlsx")
time.sleep(3)

# função para percorrer a lista e registrar
for i in range(len(dataframe)):
    descricao = dataframe.iloc[i, 0]  
    print(descricao)
    time.sleep(3)
    # inserir url de abrir chamado 
    browser.get(urlAbrirChamado)
    #area
    time.sleep(3)
    
    select_element_area = browser.find_element(By.ID,'idArea')
    select_object_area = Select(select_element_area)
    select_object_area.select_by_visible_text('Helpdesk') 
    # descrição
    time.sleep(3)
    gui.press('tab')
    time.sleep(3)
    gui.press('tab')
    time.sleep(3)
    gui.press('tab')
    time.sleep(3)
    gui.write(descricao)
    time.sleep(5)
    #departamento
    gui.press('tab')
    time.sleep(5)
    # preciso buscar Coordenadoria no select
    gui.write('Coordenadoria')
    time.sleep(3)
    gui.press('enter')
    time.sleep(3)
    gui.press('tab')
    time.sleep(3)
    gui.press('tab')
    time.sleep(3)
    gui.press('enter')
    time.sleep(5)
    print('cadastrado com sucesso!')
    time.sleep(3)

   

  





