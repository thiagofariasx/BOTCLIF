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
import sys

try:
    import functions_framework
except ImportError:
    functions_framework = None

# --- CONFIGURAÇÕES ---
URL_SISTEMA = "https://sesce.clif.rvimola.com.br"
PLANILHA_URL = "https://docs.google.com/spreadsheets/d/1_cFPlPpeFqbWR6-t8MS1XeyxCmfw2mC84ZljkMIMZFc/edit#gid=0"
DOWNLOAD_PATH = "/tmp/downloads" 

if not os.path.exists(DOWNLOAD_PATH):
    os.makedirs(DOWNLOAD_PATH)

def obter_datas_mes_atual():
    hoje = datetime.now()
    return hoje.replace(day=1).strftime("%d/%m/%Y"), hoje.replace(day=calendar.monthrange(hoje.year, hoje.month)[1]).strftime("%d/%m/%Y")

def configurar_driver():
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
    options.add_experimental_option("prefs", {"download.default_directory": DOWNLOAD_PATH})
    
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    driver.set_script_timeout(120) # Aumentado para evitar o erro de timeout
    return driver

def executar_robo():
    USUARIO = os.environ.get('USUARIO')
    SENHA = os.environ.get('SENHA')
    CHAVE_JSON_CONTENT = os.environ.get('GOOGLE_CHAVE_JSON')
    CHAVE_PATH = "/tmp/key.json"
    
    with open(CHAVE_PATH, 'w') as f: json.dump(json.loads(CHAVE_JSON_CONTENT), f)

    driver = configurar_driver()
    try:
        print("Acessando sistema...")
        driver.get(f"{URL_SISTEMA}/usuarios/login")
        time.sleep(10)
        
        # LOGIN VIA INJEÇÃO (Mais rápido e seguro)
        driver.execute_script(f"""
            document.getElementById('UsuarioLogin').value = '{USUARIO}';
            document.getElementById('UsuarioSenha').value = '{SENHA}';
            document.getElementById('UsuarioLoginForm').submit();
        """)
        time.sleep(15)
        driver.save_screenshot("pos_login.png")

        # RODAS DE DOWNLOAD
        rotinas = [
            ("Relentradaspendentes/filtroentradas", "ENTRADAS PENDENTES", "RelentradaspendenteTipofiltro"),
            ("Relsaidasgerals/filtrosaidas", "PEDIDOS EM ABERTO", "RelsaidasgeralTipofiltro")
        ]
        
        d_ini, d_fim = obter_datas_mes_atual()

        for path, aba, tipo_id in rotinas:
            print(f"Processando {aba}...")
            driver.get(f"{URL_SISTEMA}/{path}")
            time.sleep(10)
            
            # COMANDO DE CÓDIGO FONTE (Aquele que você gosta)
            driver.execute_script(f"""
                if(document.getElementById('filtro_unidade')) document.getElementById('filtro_unidade').value = '1';
                if(document.getElementById('data_inicio')) document.getElementById('data_inicio').value = '{d_ini}';
                if(document.getElementById('data_final')) document.getElementById('data_final').value = '{d_fim}';
                if(document.getElementById('{tipo_id}')) document.getElementById('{tipo_id}').value = '1';
                if(typeof EscolhaTipoRelatorio === 'function') EscolhaTipoRelatorio();
                setTimeout(function(){{ if(document.getElementById('XLSX')) document.getElementById('XLSX').click(); }}, 3000);
            """)
            
            # Espera download (simplificado)
            time.sleep(20)
            print(f"Download disparado para {aba}")
            
    except Exception as e:
        print(f"Erro: {e}")
        driver.save_screenshot("erro_fatal.png")
    finally:
        driver.quit()

if functions_framework:
    @functions_framework.http
    def main(request):
        executar_robo()
        return "Processado", 200

if __name__ == "__main__":
    executar_robo()
