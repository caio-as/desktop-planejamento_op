# Este código é o resultado do webscraping da base idc/gereisp, utilizando agendamento para startar todo dia 5:20 a.m
# Também irá renomear o arquivo e moverá para a pasta da rede Y:
# Gerando log dos eventos em arquivo .txt ('webscraping_idc_log')


from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import ElementNotInteractableException
from apscheduler.schedulers.blocking import BlockingScheduler
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import time
import shutil
import os
import logging
import sys



# Definindo suas variáveis globais
usr = "gustavo.rocha"
psw = "Rocha#Gustavo"
dt_bgn = "01/06/2023"
dt_end = "30/10/2023"



# Configurar o logger
logger = logging.getLogger('webscraping_idc_log.txt')
logger.setLevel(logging.INFO)

# Obtém o diretório onde o programa .exe está sendo executado
diretorio_do_programa = os.path.dirname(sys.argv[0])

# Criar um manipulador de arquivo para salvar o log no arquivo especificado
caminho_completo_log = os.path.join(diretorio_do_programa, 'webscraping_idc_log.txt')
file_handler = logging.FileHandler(caminho_completo_log)

# Formato das mensagens de log
formatter = logging.Formatter('%(asctime)s - %(levelname)s: %(message)s')
file_handler.setFormatter(formatter)

# Adicionar o manipulador de arquivo ao logger
logger.addHandler(file_handler)

# Redirecionar a saída padrão (stdout) para o logger
class LoggerWriter:
    def __init__(self, level):
        self.level = level

    def write(self, message):
        if message.rstrip():  # Evitar mensagens em branco
            self.level(message.rstrip())

# Redirecionar stdout e stderr para o logger
sys.stdout = LoggerWriter(logger.info)
sys.stderr = LoggerWriter(logger.error)

# Exemplo de registro de log
logger.info('Iniciando o programa de web scraping idc.')



# Função para realizar o web scraping
def do_web_scraping():

# Abre o navegador
    browser = webdriver.Chrome()
    browser.maximize_window()
    #abre o Aniel - Conect (sistema de OS da lpnet)
    browser.get("https://gerenet.idctelecom.net.br/acesso_log.php")
    time.sleep(20)

    #autentica e entra no portal
    browser.find_element(By.NAME, "login_usuario").send_keys(usr)
    browser.find_element(By.NAME, "login_senha").send_keys(psw)
    browser.find_element(By.NAME, "login_senha").send_keys(Keys.ENTER)
    time.sleep(10)
    print("Login com sucesso")

    #entre no painel gereisp
    #essa etapa do processo não vai navegar pelos botões/gereisp e menu de navegação suspenso/produção/ordens de serviço/filtro desktop matriz/base os
    browser.get("https://gerenet.idctelecom.net.br/cons_os.php?quadro=filtro_edit_listar&atualizar=sim&filtro_data=2023-03-21_07:58:18")
    time.sleep(10)
    print("Entrou no filtro desktop matriz base os")

    # Inserindo periodo do relatório
    #data inicio
    browser.find_element(By.XPATH, '/html/body/table/tbody/tr/td/form/table[2]/tbody/tr[3]/td/table[1]/tbody/tr/td[2]/table/tbody/tr[15]/td[2]/table/tbody/tr/td[1]/span/label/input').clear()
    browser.find_element(By.XPATH, '/html/body/table/tbody/tr/td/form/table[2]/tbody/tr[3]/td/table[1]/tbody/tr/td[2]/table/tbody/tr[15]/td[2]/table/tbody/tr/td[1]/span/label/input').send_keys(dt_bgn)
    time.sleep(10)
    print("Inseriu data inicial")

    #data final
    browser.find_element(By.XPATH, '/html/body/table/tbody/tr/td/form/table[2]/tbody/tr[3]/td/table[1]/tbody/tr/td[2]/table/tbody/tr[15]/td[2]/table/tbody/tr/td[3]/span/label/input').clear()
    browser.find_element(By.XPATH, '/html/body/table/tbody/tr/td/form/table[2]/tbody/tr[3]/td/table[1]/tbody/tr/td[2]/table/tbody/tr[15]/td[2]/table/tbody/tr/td[3]/span/label/input').send_keys(dt_end)
    time.sleep(10)
    print("Inseriu data final")

    #clica em gerar relatório
    browser.find_element(By.XPATH, '/html/body/table/tbody/tr/td/form/table[2]/tbody/tr[3]/td/table[3]/tbody/tr/td/input[2]').click()
    time.sleep(10)
    print("clicou em gerar relatório")

    #clica em relatório detalhado
    browser.find_element(By.XPATH, '/html/body/table/tbody/tr/td/div[1]/table/tbody/tr/td/font/b/a[1]/img').click()
    time.sleep(10)
    print("clicou em relatório detalhado")

    #clica em csv - vai exportar a planilha excel
    #time.sleep(10)
    browser.get("https://gerenet.idctelecom.net.br/cons_os.php?exporta=csv&sort_ascendente=&quadro=detalhado&atualizar=sim&sql_limite_pag_i=0&pode_fechar=sim&pag_anterior=-30&pag_posterior=30&acao=mostra_todos")
    time.sleep(20)
    print("clicou em csv")

    # Esperar um pouco para o download ser concluído
    time.sleep(60)

    # Código para renomear o arquivo, deixando padrão com a base
    # Processo para renomear e mover arquivo
    pasta_origem = r'C:\Users\caio.asouza\Downloads'
    pasta_destino = r'Y:\Operações\planejamento\_data_lake\idc_gerenet-ordens_servico'
    padrao = r'Relatório_de_Ordens_de_Serviço_'
    novo_nome = f"os_idc_2023_teste.csv"

    # Lista os arquivos na pasta de downloads
    arquivos = os.listdir(pasta_origem)

    # Filtra os arquivos que começam com o padrão esperado
    arquivos_compativeis = [arquivo for arquivo in arquivos if arquivo.startswith(padrao)]

    # Verifica se há arquivos compatíveis
    if arquivos_compativeis:
        # Seleciona o primeiro arquivo compatível (pode haver vários)
        arquivo_compativel = arquivos_compativeis[0]

        # Caminho completo para o arquivo de origem
        caminho_origem = os.path.join(pasta_origem, arquivo_compativel)

        # Move o arquivo renomeado para a pasta de destino usando shutil.move
        caminho_destino = os.path.join(pasta_destino, novo_nome)
        shutil.move(caminho_origem, caminho_destino)

        print(f"Arquivo renomeado e movido para {caminho_destino}")
        print("Programa executado com sucesso")
    else:
        print("Nenhum arquivo compatível encontrado.")

# Agendamento para executar a função todos os dias às 05:20 AM
scheduler = BlockingScheduler()
scheduler.add_job(do_web_scraping, 'cron', hour=13, minute=14)
scheduler.start()