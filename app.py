import os
import time
import glob
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# ==========================================
# CONFIGURAÇÕES E CREDENCIAIS
# ==========================================
LOGIN_USER = 'andreia.amaral@ovg.org.br'
PASSWORD_USER = '123456'
DOWNLOAD_PATH = os.path.join(os.getcwd(), "downloads")
NOME_PADRAO_ARQUIVO = "relatorio_chats_atualizado.csv" 

# Garante que a pasta existe
os.makedirs(DOWNLOAD_PATH, exist_ok=True)

# Limpeza total da pasta para evitar conflitos com arquivos antigos
for arquivo in glob.glob(os.path.join(DOWNLOAD_PATH, "*.csv")):
    try:
        os.remove(arquivo)
    except:
        pass
print("🧹 Pasta de downloads limpa para a nova extração.")

# Configurações do Navegador
chrome_options = Options()
chrome_options.add_argument('--start-maximized')
chrome_options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36')
chrome_options.add_argument('--disable-notifications') 
chrome_options.add_argument('--ignore-certificate-errors')

# 👉 BÔNUS: SE QUISER RODAR INVISÍVEL (SEM ABRIR A TELA), TIRE O '#' DA LINHA ABAIXO:
# chrome_options.add_argument('--headless=new')

chrome_options.add_experimental_option("prefs", {
    "download.default_directory": DOWNLOAD_PATH,
    "download.prompt_for_download": False,
    "safebrowsing.enabled": True 
})

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
wait = WebDriverWait(driver, 60)
actions = ActionChains(driver)

# ==========================================
# FUNÇÕES AUXILIARES
# ==========================================
def clicar_js(xpath, descricao="elemento", delay=0.5):
    """Localiza, centraliza instantaneamente e clica via JS."""
    print(f"👉 Clicando em: {descricao}")
    elemento = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
    driver.execute_script("arguments[0].scrollIntoView({behavior: 'instant', block: 'center'});", elemento)
    driver.execute_script("arguments[0].click();", elemento)
    time.sleep(delay) 

# ==========================================
# FLUXO DE AUTOMAÇÃO
# ==========================================
try:
    print("🚀 INÍCIO DA AUTOMAÇÃO: Poli Digital")
    driver.get("https://app-spa.poli.digital/login")
    
    # 1. LOGIN
    wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="root"]/div/div/div/div/div[2]/div/form/input'))).send_keys(LOGIN_USER)
    driver.find_element(By.XPATH, '//*[@id="root"]/div/div/div/div/div[2]/div/form/div[3]/div/input').send_keys(PASSWORD_USER)
    clicar_js('//*[@id="root"]/div/div/div/div/div[2]/div/form/div[6]/button', "Botão Entrar", delay=2)

    # 2. NAVEGAÇÃO
    clicar_js('//*[@id="sidebar"]/div[1]/li[5]/a/div', "Menu Relatórios", delay=2)
    print("⏳ Entrando no ambiente do Metabase...")
    iframe = wait.until(EC.presence_of_element_located((By.XPATH, "//iframe[@title='Metabase Dashboard']")))
    driver.switch_to.frame(iframe)

    # 3. FILTRAGEM
    clicar_js('//*[@id="2-T-138"]/span', "Aba Visão Geral")
    clicar_js("//button[contains(., 'Data - Período')]", "Filtro de Data")
    clicar_js("//button[contains(., 'Atual')] | //*[text()='Atual']", "Opção 'Atual'")
    clicar_js("//button[contains(., 'Mês')] | //button[text()='Mês']", "Botão 'Mês'")

    print("⏳ Aguardando o Metabase processar os dados (8s)...")
    time.sleep(8) 

    # 4. ROLAGEM E HOVER
    print("📜 Localizando a tabela 'Relatório de chats'...")
    xpath_titulo = "//*[contains(text(), 'Relátorio de chats') or contains(text(), 'Relatório de chats')]"
    titulo_tabela = wait.until(EC.presence_of_element_located((By.XPATH, xpath_titulo)))
    
    driver.execute_script("arguments[0].scrollIntoView({behavior: 'instant', block: 'center'});", titulo_tabela)
    time.sleep(2) 
    
    actions.move_to_element(titulo_tabela).perform() 
    time.sleep(1) 

    # 5. MENU OPÇÕES
    xpath_opcoes = xpath_titulo + "/ancestor::div[contains(@class, 'react-grid-item')]//button[@data-testid='public-or-embedded-dashcard-menu']"
    clicar_js(xpath_opcoes, "Botão '...' (Opções do Gráfico)")
    
    # 6. DOWNLOAD
    clicar_js("//button[@aria-label='Fazer download de resultados']", "Fazer download de resultados")
    clicar_js("//label[.//span[text()='.csv']] | //label[contains(., '.csv')]", "Formato .csv")
    
    try:
        formatacao_el = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@data-testid='keep-data-formatted']")))
        if formatacao_el.is_selected() or formatacao_el.get_attribute("checked"):
            driver.execute_script("arguments[0].click();", formatacao_el)
            print("✅ Formatação desmarcada.")
    except Exception:
        pass 

    clicar_js("//button[@data-testid='download-results-button']", "Botão 'Baixar'", delay=1)
    
    # -----------------------------------------------------------------
    # 7. MONITORAMENTO E RENOMEAÇÃO (BLINDADO CONTRA LENTIDÃO DO SERVIDOR)
    # -----------------------------------------------------------------
    print("📥 Aguardando a geração e o download do arquivo...")
    print("   (Se o relatório for muito pesado, o servidor pode levar alguns minutos)")
    
    tempo_limite = 180 # 3 MINUTOS DE PACIÊNCIA para o Metabase calcular tudo
    arquivo_baixado = None
    
    for _ in range(tempo_limite):
        arquivos_csv = glob.glob(os.path.join(DOWNLOAD_PATH, "*.csv"))
        arquivos_temp = glob.glob(os.path.join(DOWNLOAD_PATH, "*.crdownload")) # Arquivo parcial do Chrome
        
        # Só confirma quando existir o CSV e o download temporário tiver sumido
        if arquivos_csv and not arquivos_temp:
            time.sleep(2) # Pausa estratégica para o Chrome soltar o uso do arquivo no Windows
            arquivo_baixado = arquivos_csv[0]
            break
            
        time.sleep(1) 

    # Lógica de Padronização Perfeita
    if arquivo_baixado:
        caminho_final = os.path.join(DOWNLOAD_PATH, NOME_PADRAO_ARQUIVO)
        
        # Se por acaso já existir um arquivo com o nome padrão na pasta, exclui ele
        if os.path.exists(caminho_final):
            os.remove(caminho_final)
            
        os.rename(arquivo_baixado, caminho_final)
        print(f"🎉 SUCESSO ABSOLUTO! Arquivo salvo e padronizado como: {NOME_PADRAO_ARQUIVO}")
    else:
        print("⚠️ Erro: O tempo limite de 3 minutos esgotou e o download não foi concluído.")

except Exception as e:
    print(f"❌ Falha no processo: {e}")
    driver.save_screenshot("erro_log.png")
    print("📸 Screenshot de erro gerada: 'erro_log.png'")

finally:
    driver.switch_to.default_content() 
    time.sleep(2)
    driver.quit()
    print("🏁 Navegador encerrado com segurança.")