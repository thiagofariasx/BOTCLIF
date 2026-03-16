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
from datetime import datetime
import calendar
import sys

# Ajuste de encoding para evitar erros no log do GitHub
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

# --- CONFIGURAÇÕES ---
URL_SISTEMA = "https://sesce.clif.rvimola.com.br"
PLANILHA_URL = "https://docs.google.com/spreadsheets/d/1_cFPlPpeFqbWR6-t8MS1XeyxCmfw2mC84ZljkMIMZFc/edit#gid=0"

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

def configurar_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--remote-debugging-pipe")
    
    # Flags para evitar que o renderer do Chrome trave no servidor
    options.add_argument("--disable-renderer-backgrounding")
    options.add_argument("--disable-backgrounding-occluded-windows")
    options.add_argument("--disable-ipc-flooding-protection")
    options.add_argument("--disable-browser-side-navigation")
    
    # ESTRATÉGIA DE CARREGAMENTO: 'none' faz o Selenium não esperar o site carregar 100%
    # Nós controlaremos o tempo manualmente com time.sleep()
    options.page_load_strategy = 'none' 
    
    driver = webdriver.Chrome(options=options)
    
    # Timeout de 3 minutos para garantir que ele não desista fácil
    driver.set_page_load_timeout(180) 
    
    return driver

def aguardar_download(timeout=180): 
    segundos = 0
    while segundos < timeout:
        arquivos = os.listdir(DOWNLOAD_PATH)
        baixando = any(".crdownload" in f or ".tmp" in f for f in arquivos)
        finalizado = any(f.endswith((".xlsx", ".xls")) for f in arquivos)
        if finalizado and not baixando:
            time.sleep(5) # Respiro para garantir escrita do arquivo
            return True
        time.sleep(1)
        segundos += 1
    return False

def enviar_para_google(caminho_excel, nome_aba):
    try:
        print(f"Enviando dados para aba: {nome_aba}")
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
        print(f"--- SUCESSO: {nome_aba} ATUALIZADA ---")
        os.remove(caminho_excel) 
    except Exception as e:
        print(f"Erro ao enviar para Google Sheets: {e}")

def realizar_ronda(driver, wait):
    print("Iniciando rotinas de download...")
    
    # [1] Pendentes
    print("Baixando Pendentes...")
    driver.get("https://sesce.clif.rvimola.com.br/Relentradaspendentes/filtroentradas")
    time.sleep(10) # Espera carregar a página
    Select(wait.until(EC.element_to_be_clickable((By.ID, "filtro_unidade")))).select_by_value("1")
    time.sleep(3)
    driver.execute_script("document.getElementById('RelentradaspendenteTipofiltro').value = '1'; EscolhaTipoRelatorio();")
    wait.until(EC.element_to_be_clickable((By.ID, "XLSX"))).click()
    if aguardar_download(): 
        enviar_para_google(max([os.path.join(DOWNLOAD_PATH, f) for f in os.listdir(DOWNLOAD_PATH)], key=os.path.getctime), "ENTRADAS PENDENTES")

    # [2] Concluidas
    print("Baixando Concluídas...")
    d_ini, d_fim = obter_datas_mes_atual()
    driver.get("https://sesce.clif.rvimola.com.br/Relentradasconcluidasdetalhados/filtroentradas")
    time.sleep(10)
    Select(wait.until(EC.element_to_be_clickable((By.ID, "filtro_unidade")))).select_by_value("1")
    time.sleep(3)
    driver.execute_script(f"document.getElementById('data_inicio').value='{d_ini}'; document.getElementById('data_final').value='{d_fim}'; document.getElementById('RelentradasconcluidasdetalhadoTipofiltro').value='1'; EscolhaTipoRelatorio();")
    wait.until(EC.element_to_be_clickable((By.ID, "XLSX"))).click()
    if aguardar_download(): 
        enviar_para_google(max([os.path.join(DOWNLOAD_PATH, f) for f in os.listdir(DOWNLOAD_PATH)], key=os.path.getctime), "ENTRADAS CONCLUÍDAS")

    # [3] Bloqueados
    print("Baixando Bloqueados...")
    driver.get("https://sesce.clif.rvimola.com.br/Relprodutosbloqueados/listarprodutos")
    time.sleep(10)
    Select(wait.until(EC.element_to_be_clickable((By.ID, "filtro_unidade")))).select_by_value("1")
    time.sleep(3)
    driver.execute_script("document.getElementById('RelprodutosbloqueadoTipofiltro').value='1'; EscolhaTipoRelatorio();")
    wait.until(EC.element_to_be_clickable((By.ID, "XLSX"))).click()
    if aguardar_download(): 
        enviar_para_google(max([os.path.join(DOWNLOAD_PATH, f) for f in os.listdir(DOWNLOAD_PATH)], key=os.path.getctime), "BLOQUEADOS")

    # [4] Pedidos
    print("Baixando Pedidos em Aberto...")
    driver.get("https://sesce.clif.rvimola.com.br/Relsaidasgerals/filtrosaidas")
    time.sleep(15) # Site pesado
    Select(wait.until(EC.element_to_be_clickable((By.ID, "filtro_unidade")))).select_by_value("1")
    time.sleep(5)
    driver.execute_script("document.getElementById('cod_wms').value='1'; document.getElementById('filtro_nstatus_ped').value='0'; document.getElementById('RelsaidasgeralTipofiltro').value='1'; EscolhaTipoRelatorio();")
    wait.until(EC.element_to_be_clickable((By.ID, "XLSX"))).click()
    if aguardar_download(180): 
        enviar_para_google(max([os.path.join(DOWNLOAD_PATH, f) for f in os.listdir(DOWNLOAD_PATH)], key=os.path.getctime), "PEDIDOS EM ABERTO")

    # [5] Saidas
    print("Baixando Saídas Concluídas...")
    driver.get("https://sesce.clif.rvimola.com.br/Relsaidasconcluidas/listarprodutos")
    time.sleep(10)
    Select(wait.until(EC.element_to_be_clickable((By.ID, "filtro_unidade")))).select_by_value("1")
    time.sleep(3)
    driver.execute_script(f"document.getElementById('data_inicio').value = '{d_ini}'; document.getElementById('data_final').value = '{d_fim}'; document.getElementById('RelsaidasconcluidaTipofiltro').value = '1'; EscolhaTipoRelatorio();")
    wait.until(EC.element_to_be_clickable((By.ID, "XLSX"))).click()
    if aguardar_download(): 
        enviar_para_google(max([os.path.join(DOWNLOAD_PATH, f) for f in os.listdir(DOWNLOAD_PATH)], key=os.path.getctime), "SAÍDAS CONCLUÍDAS")

if __name__ == "__main__":
    # Limpeza de segurança
    for f in os.listdir(DOWNLOAD_PATH): 
        try: os.remove(os.path.join(DOWNLOAD_PATH, f))
        except: pass
        
    print(f"--- INÍCIO DA RONDA: {datetime.now()} ---")
    
    driver = None
    try:
        driver = configurar_driver()
        print("Chrome aberto. Aguardando estabilização (10s)...")
        time.sleep(10)
        
        wait = WebDriverWait(driver, 60)
        
        print(f"Acessando sistema: {URL_SISTEMA}")
        driver.get(URL_SISTEMA)
        
        print("Aguardando carregamento da página de login (30s)...")
        time.sleep(30) # Tempo vital para o modo 'none' não falhar
        
        print("Preenchendo credenciais...")
        # Usamos presence_of_element porque com o modo 'none' o elemento pode não estar "clicável" ainda
        wait.until(EC.presence_of_element_located((By.NAME, "data[Usuario][login]"))).send_keys(USUARIO)
        driver.find_element(By.NAME, "data[Usuario][senha]").send_keys(SENHA)
        driver.find_element(By.NAME, "data[Usuario][senha]").send_keys(Keys.ENTER)
        
        print("Login enviado. Aguardando entrada no painel (20s)...")
        time.sleep(20) 
        
        realizar_ronda(driver, wait)
        
    except Exception as e:
        print(f"!!! ERRO NA EXECUÇÃO !!!: {e}")
    finally:
        if driver:
            driver.quit()
        if os.path.exists(CHAVE_JSON): 
            os.remove(CHAVE_JSON)
        print(f"--- FIM DO PROCESSO: {datetime.now()} ---")
