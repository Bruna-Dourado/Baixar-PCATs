from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
from datetime import datetime
import time
import os
import shutil
import win32com.client
import re


PATH = r'C:\Users\Hiden number\Desktop\Códigos\Python\chromedriver.exe'
service = Service(PATH)
driver = webdriver.Chrome(service=service)
driver.get('https://www2.aneel.gov.br/aplicacoes_liferay/tarifa/')
time.sleep(2)
 
# Esse código é para maximizar o navegador:
driver.maximize_window()
time.sleep(3)

# Dados para Maranhão
estado = {'37': 'Equatorial MA'}

# Anos de 2014 a 2024
anos = range(2014, 2025)

for value, sigla in estado.items():
    for ano in anos:
        # Selecionar "Por Concessionária"
        driver.find_element(By.XPATH, "/html/body/table/tbody/tr[2]/td[1]/select/option[4]").click()
        time.sleep(3)
        
        # Selecionar a concessionária
        select_element = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.NAME, "Agentes"))
        )
        dropdown = Select(select_element)
        dropdown.select_by_value(value)
        time.sleep(2)
        
        # Selecionar "PCAT"
        driver.find_element(By.XPATH, "/html/body/table/tbody/tr[2]/td[3]/select/option[2]").click()
        time.sleep(2)
        
        # Selecionar o ano
        ano_select = Select(driver.find_element(By.NAME, "Ano"))
        ano_select.select_by_value(str(ano))
        time.sleep(3)
        
        # Clicar em "Pesquisar"
        driver.find_element(By.XPATH, "/html/body/table/tbody/tr[2]/td[5]/input").click()
        time.sleep(2)
        
        try:
            # Tentar baixar o arquivo
            select_element = WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, "/html/body/div/table/tbody/tr[4]/td[7]/a"))
            )
            select_element.click()
            print(f"Baixando arquivo para {sigla} - {ano}")
            time.sleep(5)
        except:
            print(f"Não foi possível baixar o arquivo para {sigla} - {ano}")
            continue

# Fechar o navegador após terminar
driver.quit()