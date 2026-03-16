import os
import time
import pandas as pd
import gspread
import json
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from datetime import datetime
import calendar
import functions_framework

# --- CONFIGURAÇÕES ---
URL_SISTEMA = "https://sesce.clif.rvimola.com.br"
PLANILHA_URL = "https://docs.google.com/spreadsheets/d/1_cFPlPpeFqbWR6-t8MS1XeyxCmfw2mC84ZljkMIMZFc/edit#gid=0"
DOWNLOAD_PATH = "/tmp/downloads" # No Google Cloud deve ser na pasta /tmp

if not os.path.exists(DOWNLOAD_PATH):
    os.makedirs(DOWNLOAD_PATH)

def configurar_driver():
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")
    # Removido Proxy pois rodaremos de São Paulo (southamerica-east1)
    
    # Instalação automática do driver no ambiente do Google
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    driver.set_page_load_timeout(120)
    return driver

# ... (Mantenha suas funções obter_datas_mes_atual, aguardar_download e enviar_para_google iguais)

@functions_framework.http
def main(request):
    USUARIO = os.environ.get('USUARIO')
    SENHA = os.environ.get('SENHA')
    CHAVE_JSON_CONTENT = os.environ.get('GOOGLE_CHAVE_JSON')
    
    # Criar arquivo temporário da chave Google
    CHAVE_JSON = "/tmp/chave_google.json"
    with open(CHAVE_JSON, 'w') as f:
        json.dump(json.loads(CHAVE_JSON_CONTENT), f)

    driver = configurar_driver()
    wait = WebDriverWait(driver, 60)
    
    try:
        print("Acessando CLIF direto de São Paulo...")
        driver.get(f"{URL_SISTEMA}/usuarios/login")
        time.sleep(15)
        
        # Sequência de Login por Teclado (ActionChains)
        actions = ActionChains(driver)
        actions.move_by_offset(500, 300).click().perform()
        time.sleep(2)
        for _ in range(3): actions.send_keys(Keys.TAB).perform()
        
        actions.send_keys(USUARIO).send_keys(Keys.TAB).send_keys(SENHA).send_keys(Keys.ENTER).perform()
        time.sleep(30)
        
        # Inicia a ronda de downloads
        # realizar_ronda(driver, wait) # Reutilize sua função de ronda aqui
        
        return "Relatórios atualizados com sucesso!", 200
    except Exception as e:
        print(f"Erro: {e}")
        return f"Falha no processo: {e}", 500
    finally:
        driver.quit()
