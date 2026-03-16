import os
import time
import pandas as pd
import gspread
import json
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from datetime import datetime
import calendar
import sys

# Ajuste de encoding para evitar erros no terminal
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

# --- CONFIGURAÇÕES GERAIS (PEGANDO DAS SECRETS DO GITHUB) ---
URL_SISTEMA = "https://sesce.clif.rvimola.com.br"
PLANILHA_URL = "https://docs.google.com/spreadsheets/d/1_cFPlPpeFqbWR6-t8MS1XeyxCmfw2mC84ZljkMIMZFc/edit#gid=0"

# Pega as informações do "Cofre" do GitHub
USUARIO = os.environ.get('USUARIO')
SENHA = os.environ.get('SENHA')
CHAVE_JSON_CONTENT = os.environ.get('GOOGLE_CHAVE_JSON')

# Cria o arquivo da chave temporariamente para o gspread usar
CHAVE_JSON = "chave_google_temp.json"
if CHAVE_JSON_CONTENT:
    with open(CHAVE_JSON, 'w') as f:
        json.dump(json.loads(CHAVE_JSON_CONTENT), f)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DOWNLOAD_PATH = os.path.join(BASE_DIR, "downloads")

if not os.path.exists(DOWNLOAD_PATH):
    os.makedirs(DOWNLOAD_PATH)

# --- FUNÇÕES DE APOIO ---

def obter_datas_mes_atual():
    hoje = datetime.now()
    data_ini = hoje.replace(day=1).strftime("%d/%m/%Y")
    ultimo_dia = calendar.monthrange(hoje.year, hoje.month)[1]
    data_fim = hoje.replace(day=ultimo_dia).strftime("%d/%m/%Y")
    return data_ini, data_fim

def configurar_driver():
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--headless") 
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--remote-debugging-port=9222") # Adicione esta linha!
    
    prefs = {
        "download.default_directory": DOWNLOAD_PATH,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    chrome_options.add_experimental_option("prefs", prefs)
    
    # Forçamos o uso do binário do Chrome que já vem no GitHub
    chrome_options.binary_location = "/usr/bin/google-chrome"
    
    return webdriver.Chrome(options=chrome_options)

def aguardar_download(timeout=60):
    segundos = 0
    while segundos < timeout:
        arquivos = os.listdir(DOWNLOAD_PATH)
        baixando = any(".crdownload" in f or ".tmp" in f for f in arquivos)
        finalizado = any(f.endswith((".xlsx", ".xls")) for f in arquivos)
        
        if finalizado and not baixando:
            time.sleep(3) 
            return True
        time.sleep(1)
        segundos += 1
    return False

def enviar_para_google(caminho_excel, nome_aba):
    try:
        print(f"Enviando dados para a aba: {nome_aba}...")
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name(CHAVE_JSON, scope)
        client = gspread.authorize(creds)
        sheet = client.open_by_url(PLANILHA_URL).worksheet(nome_aba)
        
        df = pd.read_excel(caminho_excel)

        # Tratamentos
        for col in df.columns:
            if pd.api.types.is_datetime64_any_dtype(df[col]):
                df[col] = df[col].dt.strftime('%d/%m/%Y %H:%M:%S')

        for col in df.columns:
            if df[col].dtype == 'object':
                df[col] = df[col].astype(str).str.replace("'", "").str.replace(",", ".").str.strip()
            
            serie_aux = pd.to_numeric(df[col], errors='coerce')
            if serie_aux.notnull().sum() > (len(df) * 0.5):
                df[col] = serie_aux

        df = df.fillna("")
        dados_finais = [df.columns.values.tolist()] + df.values.tolist()

        sheet.clear()
        sheet.update(dados_finais)

        horario_atual = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
        sheet.update_acell('Z1', f"Atualizado em: {horario_atual}")
        
        print(f"--- SUCESSO: Aba '{nome_aba}' atualizada! ---")
        os.remove(caminho_excel) 
    except Exception as e:
        print(f"ERRO no Google Sheets ({nome_aba}): {str(e)}")

# --- ROTINAS DE DOWNLOAD (Mesma lógica sua) ---

def rotina_pendentes(driver, wait):
    print("\n[1] Entradas Pendentes")
    driver.get("https://sesce.clif.rvimola.com.br/Relentradaspendentes/filtroentradas")
    Select(wait.until(EC.element_to_be_clickable((By.ID, "filtro_unidade")))).select_by_value("1")
    time.sleep(2)
    driver.execute_script("document.getElementById('RelentradaspendenteTipofiltro').value = '1'; EscolhaTipoRelatorio();")
    wait.until(EC.element_to_be_clickable((By.ID, "XLSX"))).click()
    if aguardar_download():
        enviar_para_google(max([os.path.join(DOWNLOAD_PATH, f) for f in os.listdir(DOWNLOAD_PATH)], key=os.path.getctime), "ENTRADAS PENDENTES")

def rotina_concluidas(driver, wait):
    print("\n[2] Entradas Concluídas")
    d_ini, d_fim = obter_datas_mes_atual()
    driver.get("https://sesce.clif.rvimola.com.br/Relentradasconcluidasdetalhados/filtroentradas")
    Select(wait.until(EC.element_to_be_clickable((By.ID, "filtro_unidade")))).select_by_value("1")
    time.sleep(2)
    driver.execute_script(f"document.getElementById('data_inicio').value='{d_ini}'; document.getElementById('data_final').value='{d_fim}'; document.getElementById('RelentradasconcluidasdetalhadoTipofiltro').value='1'; EscolhaTipoRelatorio();")
    wait.until(EC.element_to_be_clickable((By.ID, "XLSX"))).click()
    if aguardar_download():
        enviar_para_google(max([os.path.join(DOWNLOAD_PATH, f) for f in os.listdir(DOWNLOAD_PATH)], key=os.path.getctime), "ENTRADAS CONCLUÍDAS")

def rotina_bloqueados(driver, wait):
    print("\n[3] Produtos Bloqueados")
    driver.get("https://sesce.clif.rvimola.com.br/Relprodutosbloqueados/listarprodutos")
    Select(wait.until(EC.element_to_be_clickable((By.ID, "filtro_unidade")))).select_by_value("1")
    time.sleep(2)
    driver.execute_script("document.getElementById('RelprodutosbloqueadoTipofiltro').value='1'; EscolhaTipoRelatorio();")
    wait.until(EC.element_to_be_clickable((By.ID, "XLSX"))).click()
    if aguardar_download():
        enviar_para_google(max([os.path.join(DOWNLOAD_PATH, f) for f in os.listdir(DOWNLOAD_PATH)], key=os.path.getctime), "BLOQUEADOS")

def rotina_pedidos_aberto(driver, wait):
    print("\n[4] Pedidos em Aberto")
    driver.get("https://sesce.clif.rvimola.com.br/Relsaidasgerals/filtrosaidas")
    Select(wait.until(EC.element_to_be_clickable((By.ID, "filtro_unidade")))).select_by_value("1")
    time.sleep(5)
    driver.execute_script("document.getElementById('cod_wms').value='1'; document.getElementById('filtro_nstatus_ped').value='0'; document.getElementById('RelsaidasgeralTipofiltro').value='1'; EscolhaTipoRelatorio();")
    wait.until(EC.element_to_be_clickable((By.ID, "XLSX"))).click()
    if aguardar_download(90):
        enviar_para_google(max([os.path.join(DOWNLOAD_PATH, f) for f in os.listdir(DOWNLOAD_PATH)], key=os.path.getctime), "PEDIDOS EM ABERTO")

def rotina_saidas_concluidas(driver, wait):
    print("\n[5] Saídas Concluídas")
    d_ini, d_fim = obter_datas_mes_atual()
    driver.get("https://sesce.clif.rvimola.com.br/Relsaidasconcluidas/listarprodutos")
    Select(wait.until(EC.element_to_be_clickable((By.ID, "filtro_unidade")))).select_by_value("1")
    time.sleep(4)
    driver.execute_script(f"document.getElementById('data_inicio').value = '{d_ini}'; document.getElementById('data_final').value = '{d_fim}'; document.getElementById('RelsaidasconcluidaTipofiltro').value = '1'; EscolhaTipoRelatorio();")
    wait.until(EC.element_to_be_clickable((By.ID, "XLSX"))).click()
    if aguardar_download():
        enviar_para_google(max([os.path.join(DOWNLOAD_PATH, f) for f in os.listdir(DOWNLOAD_PATH)], key=os.path.getctime), "SAÍDAS CONCLUÍDAS")

# --- INICIALIZAÇÃO ---

def iniciar_ronda():
    driver = configurar_driver()
    wait = WebDriverWait(driver, 60)
    try:
        print(f"Iniciando ronda em: {datetime.now().strftime('%H:%M:%S')}")
        driver.get(URL_SISTEMA)
        
        # Login
        wait.until(EC.element_to_be_clickable((By.NAME, "data[Usuario][login]"))).send_keys(USUARIO)
        driver.find_element(By.NAME, "data[Usuario][senha]").send_keys(SENHA)
        driver.find_element(By.NAME, "data[Usuario][senha]").send_keys(Keys.ENTER)
        
        time.sleep(10)
        
        # Execução
        rotina_pendentes(driver, wait)
        rotina_concluidas(driver, wait)
        rotina_bloqueados(driver, wait)
        rotina_pedidos_aberto(driver, wait)
        rotina_saidas_concluidas(driver, wait)
        
    except Exception as e:
        print(f"ERRO NA RONDA: {e}")
    finally:
        driver.quit()
        # Limpa a chave temporária por segurança
        if os.path.exists(CHAVE_JSON):
            os.remove(CHAVE_JSON)

if __name__ == "__main__":
    # Remove lixo antes de começar
    for f in os.listdir(DOWNLOAD_PATH): 
        try: os.remove(os.path.join(DOWNLOAD_PATH, f))
        except: pass
        
    iniciar_ronda()
