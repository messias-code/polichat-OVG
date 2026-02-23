import os
import time
import glob
import pandas as pd
import numpy as np
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# ==========================================
# 1. CONFIGURAÇÕES E CAMINHOS
# ==========================================
LOGIN_USER = 'andreia.amaral@ovg.org.br'
PASSWORD_USER = '123456'

DOWNLOAD_PATH = os.path.abspath(os.path.join(os.getcwd(), "downloads"))
NOME_PADRAO_ARQUIVO = "relatorio_chats_atualizado.csv"
ARQUIVO_CSV = os.path.join(DOWNLOAD_PATH, NOME_PADRAO_ARQUIVO)
ARQUIVO_EXCEL = os.path.join(DOWNLOAD_PATH, "relatorio_chats_pronto.xlsx")

# ==========================================
# 2. FUNÇÕES AUXILIARES
# ==========================================
def formatar_tempo_exato(td):
    if pd.isna(td): return ""
    segundos_totais = max(0, int(td.total_seconds()))
    horas, resto = divmod(segundos_totais, 3600)
    minutos, segundos = divmod(resto, 60)
    return f"{horas:02d}:{minutos:02d}:{segundos:02d}"

def limpar_pasta_downloads():
    os.makedirs(DOWNLOAD_PATH, exist_ok=True)
    for extensao in ["*.csv", "*.xlsx", "*.crdownload"]:
        for arquivo in glob.glob(os.path.join(DOWNLOAD_PATH, extensao)):
            try: os.remove(arquivo)
            except: pass
    print("🧹 Pasta de downloads limpa para a nova extração.")

# ==========================================
# 3. MÓDULO DE EXTRAÇÃO INVISÍVEL (SELENIUM)
# ==========================================
def extrair_relatorio_metabase():
    print("\n" + "="*50)
    print("🚀 FASE 1: EXTRAÇÃO DE DADOS EM BACKGROUND (POLI DIGITAL)")
    print("="*50)
    
    chrome_options = Options()
    
    # 🔥 MODO FANTASMA CORRIGIDO
    chrome_options.add_argument('--headless=new')
    chrome_options.add_argument('--window-size=1920,1080') # Substitui o start-maximized
    
    chrome_options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36')
    chrome_options.add_argument('--disable-notifications') 
    chrome_options.add_argument('--ignore-certificate-errors')

    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": DOWNLOAD_PATH,
        "download.prompt_for_download": False,
        "safebrowsing.enabled": True 
    })

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    
    # Comando de baixo nível (CDP) para forçar downloads silenciosos
    driver.execute_cdp_cmd('Page.setDownloadBehavior', {
        'behavior': 'allow',
        'downloadPath': DOWNLOAD_PATH
    })
    
    wait = WebDriverWait(driver, 60)
    actions = ActionChains(driver)

    def clicar_js(xpath, descricao="elemento", delay=0.5):
        print(f"👉 A clicar em: {descricao}")
        elemento = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
        driver.execute_script("arguments[0].scrollIntoView({behavior: 'instant', block: 'center'});", elemento)
        driver.execute_script("arguments[0].click();", elemento)
        time.sleep(delay) 

    sucesso = False

    try:
        driver.get("https://app-spa.poli.digital/login")
        
        # 1. LOGIN
        wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="root"]/div/div/div/div/div[2]/div/form/input'))).send_keys(LOGIN_USER)
        driver.find_element(By.XPATH, '//*[@id="root"]/div/div/div/div/div[2]/div/form/div[3]/div/input').send_keys(PASSWORD_USER)
        clicar_js('//*[@id="root"]/div/div/div/div/div[2]/div/form/div[6]/button', "Botão Entrar", delay=2)

        # 2. NAVEGAÇÃO
        clicar_js('//*[@id="sidebar"]/div[1]/li[5]/a/div', "Menu Relatórios", delay=2)
        print("⏳ A entrar no ambiente do Metabase (Invisível)...")
        iframe = wait.until(EC.presence_of_element_located((By.XPATH, "//iframe[@title='Metabase Dashboard']")))
        driver.switch_to.frame(iframe)

        # 3. FILTRAGEM (ANO INTEIRO)
        clicar_js('//*[@id="2-T-138"]/span', "Aba Visão Geral")
        clicar_js("//button[contains(., 'Data - Período')]", "Filtro de Data")
        clicar_js("//button[contains(., 'Atual')] | //*[text()='Atual']", "Opção 'Atual'")
        clicar_js("//button[.//span[text()='Ano']] | //button[contains(., 'Ano')]", "Botão 'Ano'")

        print("⏳ A aguardar o Metabase processar os dados de UM ANO (15s)...")
        time.sleep(15) 

        # 4. ROLAGEM E HOVER
        print("📜 A localizar a tabela 'Relatório de chats'...")
        xpath_titulo = "//*[contains(text(), 'Relátorio de chats') or contains(text(), 'Relatório de chats')]"
        titulo_tabela = wait.until(EC.presence_of_element_located((By.XPATH, xpath_titulo)))
        
        driver.execute_script("arguments[0].scrollIntoView({behavior: 'instant', block: 'center'});", titulo_tabela)
        time.sleep(2) 
        actions.move_to_element(titulo_tabela).perform() 
        time.sleep(1) 

        # 5. MENU OPÇÕES E DOWNLOAD
        xpath_opcoes = xpath_titulo + "/ancestor::div[contains(@class, 'react-grid-item')]//button[@data-testid='public-or-embedded-dashcard-menu']"
        clicar_js(xpath_opcoes, "Botão '...' (Opções do Gráfico)")
        
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
        
        # 6. MONITORIZAÇÃO DE DOWNLOAD (5 MINUTOS)
        print("📥 A aguardar o download em background...")
        tempo_limite = 300 
        arquivo_baixado = None
        
        for _ in range(tempo_limite):
            arquivos_csv = glob.glob(os.path.join(DOWNLOAD_PATH, "*.csv"))
            arquivos_temp = glob.glob(os.path.join(DOWNLOAD_PATH, "*.crdownload"))
            
            if arquivos_csv and not arquivos_temp:
                time.sleep(2) 
                arquivo_baixado = arquivos_csv[0]
                break
                
            time.sleep(1) 

        if arquivo_baixado:
            if os.path.exists(ARQUIVO_CSV): os.remove(ARQUIVO_CSV)
            os.rename(arquivo_baixado, ARQUIVO_CSV)
            print(f"🎉 DOWNLOAD CONCLUÍDO! Ficheiro salvo como: {NOME_PADRAO_ARQUIVO}")
            sucesso = True
        else:
            print("⚠️ Erro: Tempo limite esgotado. O download não foi concluído.")

    except Exception as e:
        print(f"❌ Falha no processo de extração: {e}")
        driver.save_screenshot("erro_log_extracao.png")

    finally:
        driver.switch_to.default_content() 
        time.sleep(2)
        driver.quit()
        print("🏁 Navegador encerrado.")
        return sucesso

# ==========================================
# 4. MÓDULO DE TRATAMENTO DE DADOS (PANDAS)
# ==========================================
def analisar_e_limpar_dados():
    print("\n" + "="*50)
    print("📊 FASE 2: ENGENHARIA DE DADOS E FORMATAÇÃO EXCEL")
    print("="*50)

    try:
        df = pd.read_csv(ARQUIVO_CSV, sep=',', low_memory=False)

        colunas_texto = ['Telefone do contato', 'Id do atendimento', 'Id do cliente', 'CPF do contato']
        for col in colunas_texto:
            if col in df.columns:
                df[col] = df[col].fillna(-1).astype(str).str.replace(r'\.0$', '', regex=True).replace('-1', '')

        dt_chegada = pd.to_datetime(df.get('Data de criação do chat'), errors='coerce').dt.tz_localize(None)
        dt_resposta = pd.to_datetime(df.get('Data de primeira resposta'), errors='coerce').dt.tz_localize(None)
        dt_fim = pd.to_datetime(df.get('Data de finalização do chat'), errors='coerce').dt.tz_localize(None)

        horas = dt_chegada.dt.hour
        df['Período do Dia'] = np.select(
            [(horas >= 0) & (horas < 6), (horas >= 6) & (horas < 12), (horas >= 12) & (horas < 18), (horas >= 18) & (horas <= 23)],
            ['Madrugada', 'Manhã', 'Tarde', 'Noite'], default='Desconhecido'
        )

        def verificar_expediente(dt):
            if pd.isna(dt): return "Desconhecido"
            if dt.dayofweek >= 5: return "NÃO (Fim de Semana)" 
            minutos_do_dia = dt.hour * 60 + dt.minute
            if (8 * 60 + 1) <= minutos_do_dia <= (17 * 60 + 59): return "SIM"
            return "NÃO (Fora do Horário)"
            
        df['Dentro do Expediente?'] = dt_chegada.apply(verificar_expediente)

        def avaliar_espera(row, data_chegada, data_resposta, data_fim):
            if pd.isna(data_resposta): 
                if pd.isna(data_fim): return "⏳ Na Fila (Ainda em Aberto)"
                delta_vacuo = (data_fim - data_chegada).total_seconds()
                if data_fim.date() > data_chegada.date(): return "⚠️ Vácuo até o Dia Seguinte (Fechado sem resposta)"
                minutos_vacuo = delta_vacuo / 60
                if pd.isna(row.get('Atendente')) or str(row.get('Atendente')).strip() == '':
                    return f"🤖 Sistema/Robô (Encerrado após {int(minutos_vacuo)} min)"
                else: return f"👻 Vácuo Total (Fechado após {int(minutos_vacuo)} min)"

            delta = (data_resposta - data_chegada).total_seconds()
            if data_resposta.date() > data_chegada.date(): return "⚠️ Passou para o Dia Seguinte"
            
            minutos = delta / 60
            if minutos <= 5: return "🟢 Rápido (< 5 min)"
            elif minutos <= 15: return "🟡 Aceitável (5 a 15 min)"
            else: return "🟠 Demorado (> 15 min)"
            
        df['Avaliação da Espera'] = [avaliar_espera(row, c, r, f) for row, c, r, f in zip(df.to_dict('records'), dt_chegada, dt_resposta, dt_fim)]

        def diagnosticar_conversa(row, resp, fim):
            if pd.isna(row.get('Atendente')): return "🤖 Retido no Robô"
            if pd.isna(resp): return "👻 Ignorado (Atendente nunca respondeu)"
            if pd.notna(fim):
                if (fim - resp).total_seconds() < 60: return "⚡ Fechamento Imediato (Sem diálogo longo)"
                else: return "✅ Atendimento com Interação"
            return "⏳ Em Andamento"

        df['Diagnóstico da Conversa'] = [diagnosticar_conversa(row, r, f) for row, r, f in zip(df.to_dict('records'), dt_resposta, dt_fim)]
        df['Status Final'] = np.where(dt_fim.notna(), "Encerrado", "Em Aberto")

        data_limite_espera = dt_resposta.fillna(dt_fim)
        df['Tempo de Espera (Fila)'] = (data_limite_espera - dt_chegada).apply(formatar_tempo_exato)
        df['Tempo de Conversa (Atendimento)'] = np.where(dt_resposta.notna(), (dt_fim - dt_resposta).apply(formatar_tempo_exato), "")
        df['Tempo Total (Início ao Fim)'] = (dt_fim - dt_chegada).apply(formatar_tempo_exato)

        df['1. Data de Entrada'] = dt_chegada.dt.normalize()
        df['1. Hora de Entrada'] = dt_chegada.dt.time
        df['2. Data da 1ª Resposta'] = dt_resposta.dt.normalize()
        df['2. Hora da 1ª Resposta'] = dt_resposta.dt.time
        df['3. Data de Encerramento'] = dt_fim.dt.normalize()
        df['3. Hora de Encerramento'] = dt_fim.dt.time

        ordem_historia = [
            'Id do atendimento', 'Cliente', 'Telefone do contato',
            '1. Data de Entrada', '1. Hora de Entrada', 'Período do Dia', 'Dentro do Expediente?',
            'Houve redirecionamento', 'Departamento do Chat', 'Atendente', 'Tempo de Espera (Fila)', 'Avaliação da Espera',
            '2. Data da 1ª Resposta', '2. Hora da 1ª Resposta',
            'Diagnóstico da Conversa', 'Tempo de Conversa (Atendimento)',
            '3. Data de Encerramento', '3. Hora de Encerramento', 'Fechado por', 'Tempo Total (Início ao Fim)', 'Status Final',
            'Motivo do serviço', 'Motivo do fechamento'
        ]

        colunas_finais = [col for col in ordem_historia if col in df.columns]
        df_final = df[colunas_finais].copy()

        print("🎨 A criar o ficheiro Excel Premium...")
        writer = pd.ExcelWriter(ARQUIVO_EXCEL, engine='xlsxwriter', datetime_format='dd/mm/yyyy', engine_kwargs={'options': {'strings_to_urls': False}})
        nome_aba = 'Relatorio_Chats'
        df_final.to_excel(writer, index=False, header=False, startrow=1, sheet_name=nome_aba)
        
        workbook = writer.book
        ws = writer.sheets[nome_aba]
        ws.set_tab_color('#FF8C00') 
        
        (max_r, max_c) = df_final.shape
        if max_r > 0:
            ws.add_table(0, 0, max_r, max_c - 1, {
                'columns': [{'header': str(c)} for c in df_final.columns],
                'style': 'Table Style Medium 9', 'name': 'Tab_Chats'
            })
        else:
            ws.write_row(0, 0, df_final.columns)

        ws.ignore_errors({'number_stored_as_text': 'A1:XFD1048576'})

        fmt_hora = workbook.add_format({'num_format': 'hh:mm:ss'})
        fmt_central = workbook.add_format({'align': 'center'})

        for i, col in enumerate(df_final.columns):
            try: tamanho = int(df_final[col].fillna("").astype(str).str.len().max())
            except: tamanho = 10
            largura = min(max(tamanho, len(str(col))) + 2, 45)

            if "Hora " in col or "Tempo " in col: ws.set_column(i, i, largura, fmt_hora)
            elif "Houve" in col: ws.set_column(i, i, largura, fmt_central)
            else: ws.set_column(i, i, largura)

        writer.close()
        print(f"🏆 SUCESSO TOTAL! O Pipeline terminou. Ficheiro final: {ARQUIVO_EXCEL}")

    except Exception as e:
        print(f"❌ Ocorreu um erro no tratamento de dados: {e}")

# ==========================================
# 5. ORQUESTRADOR PRINCIPAL (PIPELINE)
# ==========================================
if __name__ == "__main__":
    limpar_pasta_downloads()
    
    sucesso_download = extrair_relatorio_metabase()
    
    if sucesso_download:
        analisar_e_limpar_dados()
    else:
        print("\n⚠️ O tratamento de dados foi cancelado porque o ficheiro CSV não pôde ser descarregado.")