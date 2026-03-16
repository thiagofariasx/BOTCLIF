import os
import time
import pandas as pd
import gspread
import json
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from datetime import datetime
import calendar
import sys

# Ajuste de encoding
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

# --- CONFIGURAÇÕES ---
URL_SISTEMA = "https://sesce.clif.rvimola.com.br"
PLANILHA_URL = "https://docs.google.com/spreadsheets/d/1_cFPlPpeFqbWR6-t8MS1XeyxCmfw2mC84ZljkMIMZFc/edit#gid=0"

# LISTA DE PROXIES BRASILEIROS (Baseado na sua imagem)
# Adicionei os mais rápidos da sua lista para rotacionar
LISTA_PROXIES = [
    "186.235.123.3:8080",
    "45.168.49.2:8080",
    "179.185.89.251:8080",
    "186.235.201.10:8080",
    "186.192.78.5:8080"
]

USUARIO = os.environ.get('USUARIO')
SENHA = os.environ.get('SENHA')
CHAVE_JSON_CONTENT = os.environ.get('GOOGLE_CHAVE_JSON')

CHAVE_JSON = "chave_google_temp.json"
if CHAVE_JSON_CONTENT:
    with open(CHAVE_JSON, 'w') as f:
        json.dump(json.loads(CHAVE_JSON_CONTENT), f)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DOWNLOAD_PATH = os.path.join(BASE_DIR, "downloads")
if not os.path.exists(DOWNLOAD_PATH):
    os.makedirs(DOWNLOAD_PATH)

def obter_datas_mes_atual():
    hoje = datetime.now()
    data_ini = hoje.replace(day=1).strftime("%d/%m/%Y")
    ultimo_dia = calendar.monthrange(hoje.year, hoje.month)[1]
    data_fim = hoje.replace(day=ultimo_dia).strftime("%d/%m/%Y")
    return data_ini, data_fim

def configurar_driver(proxy_atual=None):
    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    
    if proxy_atual:
        print(f"Configurando Proxy Brasileiro: {proxy_atual}")
        options.add_argument(f'--proxy-server={proxy_atual}')
        
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36")
    options.page_load_strategy = 'none' 
    
    driver = webdriver.Chrome(options=options)
    driver.set_page_load_timeout(120)
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

def enviar_para_google(caminho_excel, nome_aba):
    try:
        print(f"Enviando dados para: {nome_aba}")
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name(CHAVE_JSON, scope)
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
        sheet.update_acell('Z1', f"Atualizado: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
        print(f"--- SUCESSO: {nome_aba} ---")
        os.remove(caminho_excel) 
    except Exception as e:
        print(f"Erro Google Sheets: {e}")

def realizar_ronda(driver, wait):
    print("Iniciando rotinas de download...")
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
            driver.get(f"https://sesce.clif.rvimola.com.br/{path}")
            time.sleep(25)
            Select(wait.until(EC.presence_of_element_located((By.ID, "filtro_unidade")))).select_by_value("1")
            
            if "concluidas" in path:
                driver.execute_script(f"document.getElementById('data_inicio').value='{d_ini}'; document.getElementById('data_final').value='{d_fim}';")
            
            driver.execute_script(f"document.getElementById('{tipo_filtro_id}').value = '1'; EscolhaTipoRelatorio();")
            time.sleep(5)
            driver.execute_script("document.getElementById('XLSX').click();")
            
            if aguardar_download(200 if "saidasgerals" in path else 120):
                enviar_para_google(max([os.path.join(DOWNLOAD_PATH, f) for f in os.listdir(DOWNLOAD_PATH)], key=os.path.getctime), aba)
        except Exception as e:
            print(f"Erro em {aba}: {e}")

if __name__ == "__main__":
    for f in os.listdir(DOWNLOAD_PATH): 
        try: os.remove(os.path.join(DOWNLOAD_PATH, f))
        except: pass
        
    print(f"--- INÍCIO DA RONDA: {datetime.now()} ---")
    
    driver = None
    logado = False

    # TENTA CADA PROXY DA LISTA ATÉ UM FUNCIONAR
    for proxy in LISTA_PROXIES:
        try:
            if driver: driver.quit()
            driver = configurar_driver(proxy)
            
            print(f"Testando conexão com o CLIF via {proxy}...")
            driver.get("https://sesce.clif.rvimola.com.br/usuarios/login")
            time.sleep(30)
            
            # Se o título da página aparecer, o proxy funcionou
            if "CLIF" in driver.title:
                print("Conexão estabelecida! Iniciando Login...")
                
                actions = ActionChains(driver)
                actions.move_by_offset(500, 300).click().perform()
                time.sleep(2)
                for _ in range(3): actions.send_keys(Keys.TAB).perform()
                
                actions.send_keys(USUARIO).send_keys(Keys.TAB).send_keys(SENHA).send_keys(Keys.ENTER).perform()
                time.sleep(40)
                
                print(f"Página após login: {driver.current_url}")
                realizar_ronda(driver, WebDriverWait(driver, 60))
                logado = True
                break # Sai do loop de proxies se deu certo
            else:
                print(f"Proxy {proxy} falhou em carregar o site (Título não encontrado).")
                
        except Exception as e:
            print(f"Erro com o proxy {proxy}: {e}")
            continue

    if not logado:
        print("!!! TODOS OS PROXIES FALHARAM !!!")
        if driver: driver.save_screenshot("erro_total.png")
    
    if driver: driver.quit()
    if os.path.exists(CHAVE_JSON): os.remove(CHAVE_JSON)
    print(f"--- FIM DO PROCESSO: {datetime.now()} ---")
