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

# Tenta importar o framework do Google, se falhar (no GitHub), segue normal
try:
    import functions_framework
except ImportError:
    functions_framework = None

# Ajuste de encoding
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

# --- CONFIGURAÇÕES ---
URL_SISTEMA = "https://sesce.clif.rvimola.com.br"
PLANILHA_URL = "https://docs.google.com/spreadsheets/d/1_cFPlPpeFqbWR6-t8MS1XeyxCmfw2mC84ZljkMIMZFc/edit#gid=0"
DOWNLOAD_PATH = "/tmp/downloads" 

if not os.path.exists(DOWNLOAD_PATH):
    os.makedirs(DOWNLOAD_PATH)

def obter_datas_mes_atual():
    hoje = datetime.now()
    data_ini = hoje.replace(day=1).strftime("%d/%m/%Y")
    ultimo_dia = calendar.monthrange(hoje.year, hoje.month)[1]
    data_fim = hoje.replace(day=ultimo_dia).strftime("%d/%m/%Y")
    return data_ini, data_fim

def configurar_driver():
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36")
    
    prefs = {"download.default_directory": DOWNLOAD_PATH}
    options.add_experimental_option("prefs", prefs)
    options.page_load_strategy = 'none'
    
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    driver.set_page_load_timeout(180)
    return driver

def aguardar_download(timeout=180):
    segundos = 0
    while segundos < timeout:
        arquivos = os.listdir(DOWNLOAD_PATH)
        baixando = any(".crdownload" in f or ".tmp" in f for f in arquivos)
        finalizado = any(f.endswith((".xlsx", ".xls")) for f in arquivos)
        if finalizado and not baixando:
            time.sleep(5)
            return True
        time.sleep(1)
        segundos += 1
    return False

def enviar_para_google(caminho_excel, nome_aba, chave_path):
    try:
        print(f"Enviando dados para aba: {nome_aba}")
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name(chave_path, scope)
        client = gspread.authorize(creds)
        sheet = client.open_by_url(PLANILHA_URL).worksheet(nome_aba)
        
        df = pd.read_excel(caminho_excel)
        for col in df.columns:
            if pd.api.types.is_datetime64_any_dtype(df[col]):
                df[col] = df[col].dt.strftime('%d/%m/%Y %H:%M:%S')
            if df[col].dtype == 'object':
                df[col] = df[col].astype(str).str.replace("'", "").str.strip()
        
        df = df.fillna("")
        dados = [df.columns.values.tolist()] + df.values.tolist()
        
        sheet.clear()
        sheet.update(dados)
        sheet.update_acell('Z1', f"Atualizado (SP): {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
        os.remove(caminho_excel)
    except Exception as e:
        print(f"Erro Google Sheets na aba {nome_aba}: {e}")

def realizar_ronda(driver, wait, chave_path):
    print("Iniciando rotinas de download de relatórios...")
    rotinas = [
        ("Relentradaspendentes/filtroentradas", "ENTRADAS PENDENTES", "RelentradaspendenteTipofiltro"),
        ("Relentradasconcluidasdetalhados/filtroentradas", "ENTRADAS CONCLUÍDAS", "RelentradasconcluidasdetalhadoTipofiltro"),
        ("Relprodutosbloqueados/listarprodutos", "BLOQUEADOS", "RelprodutosbloqueadoTipofiltro"),
        ("Relsaidasgerals/filtrosaidas", "PEDIDOS EM ABERTO", "RelsaidasgeralTipofiltro"),
        ("Relsaidasconcluidas/listarprodutos", "SAÍDAS CONCLUÍDAS", "RelsaidasconcluidaTipofiltro")
    ]
    d_ini, d_fim = obter_datas_mes_atual()

    for path, aba, tipo_filtro_id in rotinas:
        try:
            print(f"Processando {aba}...")
            driver.get(f"{URL_SISTEMA}/{path}")
            time.sleep(25)
            
            # Unidade Operacional
            Select(wait.until(EC.presence_of_element_located((By.ID, "filtro_unidade")))).select_by_value("1")
            
            if "concluidas" in path:
                driver.execute_script(f"document.getElementById('data_inicio').value='{d_ini}'; document.getElementById('data_final').value='{d_fim}';")
            if "saidasgerals" in path:
                driver.execute_script("document.getElementById('cod_wms').value='1'; document.getElementById('filtro_nstatus_ped').value='0';")
                time.sleep(5)

            driver.execute_script(f"document.getElementById('{tipo_filtro_id}').value = '1'; EscolhaTipoRelatorio();")
            time.sleep(5)
            driver.execute_script("document.getElementById('XLSX').click();")
            
            if aguardar_download(200):
                arquivo_recente = max([os.path.join(DOWNLOAD_PATH, f) for f in os.listdir(DOWNLOAD_PATH)], key=os.path.getctime)
                enviar_para_google(arquivo_recente, aba, chave_path)
        except Exception as e:
            print(f"Erro na aba {aba}: {e}")

def executar_robo():
    USUARIO = os.environ.get('USUARIO')
    SENHA = os.environ.get('SENHA')
    CHAVE_JSON_CONTENT = os.environ.get('GOOGLE_CHAVE_JSON')
    
    CHAVE_TEMP_PATH = "/tmp/google_key_clif.json"
    if CHAVE_JSON_CONTENT:
        with open(CHAVE_TEMP_PATH, 'w') as f:
            json.dump(json.loads(CHAVE_JSON_CONTENT), f)

    print(f"--- INÍCIO DA RONDA (SÃO PAULO): {datetime.now()} ---")
    
    # Limpa a pasta temporária antes de começar
    for f in os.listdir(DOWNLOAD_PATH):
        try: os.remove(os.path.join(DOWNLOAD_PATH, f))
        except: pass

    driver = configurar_driver()
    wait = WebDriverWait(driver, 60)
    
    try:
        print("Abrindo tela de login...")
        driver.get(f"{URL_SISTEMA}/usuarios/login")
        time.sleep(35)

        print("Simulando login por teclado...")
        actions = ActionChains(driver)
        actions.move_by_offset(500, 300).click().perform()
        time.sleep(2)
        
        for _ in range(3):
            actions.send_keys(Keys.TAB).perform()
            time.sleep(0.5)
            
        actions.send_keys(USUARIO).send_keys(Keys.TAB).send_keys(SENHA).send_keys(Keys.ENTER).perform()
        print("Login enviado. Aguardando estabilização...")
        time.sleep(45) 
        
        realizar_ronda(driver, wait, CHAVE_TEMP_PATH)
        print("--- SUCESSO TOTAL ---")
        
    except Exception as e:
        print(f"!!! ERRO NO PROCESSO !!!: {e}")
    finally:
        driver.quit()
        if os.path.exists(CHAVE_TEMP_PATH):
            os.remove(CHAVE_TEMP_PATH)

# Wrapper para o Cloud Functions
if functions_framework:
    @functions_framework.http
    def main(request):
        executar_robo()
        return "OK", 200

# Execução manual ou GitHub
if __name__ == "__main__":
    executar_robo()
