import os
import time
import pandas as pd
import gspread
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

# --- CONFIGURAÇÕES GERAIS ---
URL_SISTEMA = "https://sesce.clif.rvimola.com.br"
PLANILHA_URL = "https://docs.google.com/spreadsheets/d/1_cFPlPpeFqbWR6-t8MS1XeyxCmfw2mC84ZljkMIMZFc/edit#gid=0"
USUARIO = "THIAGO.FARIAS"
SENHA = "251090#"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DOWNLOAD_PATH = os.path.join(BASE_DIR, "downloads")
CHAVE_JSON = os.path.join(BASE_DIR, "chave_google.json")

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
    # LINHAS NECESSÁRIAS PARA RODAR ONLINE:
    chrome_options.add_argument("--headless") # Roda sem abrir janela
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    
    prefs = {
        "download.default_directory": DOWNLOAD_PATH,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    chrome_options.add_experimental_option("prefs", prefs)
    # No GitHub, ele já tem o driver instalado, então o código fica assim:
    return webdriver.Chrome(options=chrome_options)

def aguardar_download(timeout=60):
    """Espera o download terminar e garante que não seja arquivo temporário"""
    segundos = 0
    while segundos < timeout:
        arquivos = os.listdir(DOWNLOAD_PATH)
        # Verifica se tem arquivos terminando em .crdownload ou .tmp (ainda baixando)
        baixando = any(".crdownload" in f or ".tmp" in f for f in arquivos)
        finalizado = any(f.endswith((".xlsx", ".xls")) for f in arquivos)
        
        if finalizado and not baixando:
            time.sleep(2) # Respiro final para o Windows liberar o arquivo
            return True
        time.sleep(1)
        segundos += 1
    return False

def enviar_para_google(caminho_excel, nome_aba):
    try:
        print(f"Processando e enviando para aba: {nome_aba}...")
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name(CHAVE_JSON, scope)
        client = gspread.authorize(creds)
        sheet = client.open_by_url(PLANILHA_URL).worksheet(nome_aba)
        
        df = pd.read_excel(caminho_excel)

        # 1. TRATAMENTO DE DATAS
        for col in df.columns:
            if pd.api.types.is_datetime64_any_dtype(df[col]):
                df[col] = df[col].dt.strftime('%d/%m/%Y %H:%M:%S')

        # 2. LIMPEZA DE NÚMEROS E APÓSTROFOS
        for col in df.columns:
            if df[col].dtype == 'object':
                df[col] = df[col].astype(str).str.replace("'", "").str.replace(",", ".").str.strip()
            
            serie_aux = pd.to_numeric(df[col], errors='coerce')
            if serie_aux.notnull().sum() > (len(df) * 0.5):
                df[col] = serie_aux

        df = df.fillna("")
        
        # 3. PREPARAÇÃO DOS DADOS (Lista de Listas)
        dados_finais = [df.columns.values.tolist()]
        for row in df.values.tolist():
            processed_row = [str(item) if isinstance(item, (datetime, pd.Timestamp)) else item for item in row]
            dados_finais.append(processed_row)

        # Ação no Google Sheets
        sheet.clear()
        sheet.update(dados_finais)

        horario_atual = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
        sheet.update_acell('Z1', f"Atualizado em: {horario_atual}")
        
        print(f"--- SUCESSO: Aba '{nome_aba}' atualizada! ---")
        os.remove(caminho_excel) 
    except Exception as e:
        print(f"ERRO no Google Sheets ({nome_aba}): {str(e)}")

# --- ROTINAS DE DOWNLOAD ---

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
    # Para este relatório que é maior, esperamos até 90 segundos
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
    wait = WebDriverWait(driver, 45)
    try:
        print("Acessando CLIF para Login...")
        driver.get(URL_SISTEMA)
        
        try:
            btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(., 'Entrar')] | //button[contains(., 'Entrar')]")))
            driver.execute_script("arguments[0].click();", btn)
        except: pass

        wait.until(EC.element_to_be_clickable((By.NAME, "data[Usuario][login]"))).send_keys(USUARIO)
        driver.find_element(By.NAME, "data[Usuario][senha]").send_keys(SENHA)
        driver.find_element(By.NAME, "data[Usuario][senha]").send_keys(Keys.ENTER)
        
        time.sleep(8)
        driver.get("https://sesce.clif.rvimola.com.br/Homes/index/all")
        
        # Execução das rotinas
        rotina_pendentes(driver, wait)
        rotina_concluidas(driver, wait)
        rotina_bloqueados(driver, wait)
        rotina_pedidos_aberto(driver, wait)
        rotina_saidas_concluidas(driver, wait)
        
    except Exception as e:
        print(f"ERRO NA RONDA: {e}")
    finally:
        driver.quit()

if __name__ == "__main__":
    while True:
        print(f"\n===== INÍCIO RONDA GERAL (5 RELATÓRIOS): {time.strftime('%H:%M:%S')} =====")
        # Limpa downloads antigos antes de começar a nova ronda
        for f in os.listdir(DOWNLOAD_PATH): 
            try: os.remove(os.path.join(DOWNLOAD_PATH, f))
            except: pass
            
        iniciar_ronda()
        print(f"===== FIM DA RONDA. PRÓXIMA EM 20 MINUTOS =====")
        time.sleep(1200)
        