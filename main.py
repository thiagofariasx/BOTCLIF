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

# Tenta importar o framework do Google para não dar erro no GitHub
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
    driver.set_page_load_timeout(120)
    return driver

def aguardar_download(timeout=120):
    segundos = 0
    while segundos < timeout:
        arquivos = os.listdir(DOWNLOAD_PATH)
        if any(f.endswith((".xlsx", ".xls")) for f in arquivos) and not any(".crdownload" in f for f in arquivos):
            time.sleep(3)
            return True
        time.sleep(1)
        segundos += 1
    return False

def enviar_para_google(caminho_excel, nome_aba, chave_path):
    try:
        print(f"Enviando dados: {nome_aba}")
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name(chave_path, scope)
        client = gspread.authorize(creds)
        sheet = client.open_by_url(PLANILHA_URL).worksheet(nome_aba)
        
        df = pd.read_excel(caminho_excel).fillna("")
        for col in df.columns:
            if df[col].dtype == 'object':
                df[col] = df[col].astype(str).str.replace("'", "").str.strip()
        
        dados = [df.columns.values.tolist()] + df.values.tolist()
        sheet.clear()
        sheet.update(dados)
        os.remove(caminho_excel)
        print(f"--- {nome_aba} OK ---")
    except Exception as e:
        print(f"Erro Sheets {nome_aba}: {e}")

def realizar_ronda(driver, chave_path):
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
            print(f"Acessando {aba}...")
            driver.get(f"{URL_SISTEMA}/{path}")
            time.sleep(15)
            
            # Injeção JS pura (Sem $)
            driver.execute_script(f"""
                var u = document.getElementById('filtro_unidade'); if(u) u.value = '1';
                var i = document.getElementById('data_inicio'); if(i) i.value = '{d_ini}';
                var f = document.getElementById('data_final'); if(f) f.value = '{d_fim}';
                var t = document.getElementById('{tipo_filtro_id}'); if(t) t.value = '1';
                if(typeof EscolhaTipoRelatorio === 'function') EscolhaTipoRelatorio();
                setTimeout(function(){{ var b = document.getElementById('XLSX'); if(b) b.click(); }}, 3000);
            """)
            
            if aguardar_download():
                arq = max([os.path.join(DOWNLOAD_PATH, f) for f in os.listdir(DOWNLOAD_PATH)], key=os.path.getctime)
                enviar_para_google(arq, aba, chave_path)
        except Exception as e:
            print(f"Erro em {aba}: {e}")

def executar_robo():
    USUARIO = os.environ.get('USUARIO')
    SENHA = os.environ.get('SENHA')
    CHAVE_JSON_CONTENT = os.environ.get('GOOGLE_CHAVE_JSON')
    CHAVE_PATH = "/tmp/key.json"
    
    if not CHAVE_JSON_CONTENT: return

    with open(CHAVE_PATH, 'w') as f: json.dump(json.loads(CHAVE_JSON_CONTENT), f)

    driver = configurar_driver()
    try:
        print("Abrindo login via São Paulo...")
        driver.get(f"{URL_SISTEMA}/usuarios/login")
        time.sleep(15)
        
        driver.execute_script(f"""
            document.getElementById('UsuarioLogin').value = '{USUARIO}';
            document.getElementById('UsuarioSenha').value = '{SENHA}';
            document.getElementById('UsuarioLoginForm').submit();
        """)
        time.sleep(20)
        
        if "login" not in driver.current_url:
            realizar_ronda(driver, CHAVE_PATH)
        else:
            print("Ainda na tela de login. Verifique credenciais.")
    finally:
        driver.quit()

if functions_framework:
    @functions_framework.http
    def main(request):
        executar_robo()
        return "Concluído", 200

if __name__ == "__main__":
    executar_robo()
