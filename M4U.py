import time
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime

now = datetime.now()

user = "email"
passw = "senha"
dt1 = now.strftime('%d/%m/%Y') #DateTime.Now.ToString("dd/MM/yyyy")
dt2 = now.strftime('%d/%m/%Y') #DateTime.Now.ToString("dd/MM/yyyy")
rel1 = "https://portaloi.m4u.com.br/ExibeReport.aspx?id=1870"
rel2 = "https://portaloi.m4u.com.br/ListaArquivosS3.aspx?key=Oi/OI_ControleDigital/OUT/AdesoesDetalhadoHoraHora/"
portal = "https://portaloi.m4u.com.br/Login.aspx"
downloadFilepath = "\\\\netprd03\\PLANEJAMENTO_PERFORMANCE_MG_BA\\OI\\Relatorios\\Parcial_Controle\\Base"
#downloadFilepath = "C:\\Users\\Administrador\\Desktop\\JUNIOR\\Py\\M4U"

opt = webdriver.ChromeOptions()
prefs = {"download.default_directory" : downloadFilepath}
opt.add_experimental_option("prefs",prefs)

driver = webdriver.Chrome("chromedriver.exe", chrome_options=opt)

driver.get(portal)
wait = WebDriverWait(driver, 1000)
wait.until(EC.element_to_be_clickable((By.ID, 'ctl00_MainContent_LoginButton')))
driver.find_element_by_xpath("/html/body/div[1]/form/div[3]/input[1]").send_keys(user)
driver.find_element_by_xpath("/html/body/div[1]/form/div[3]/input[2]").send_keys(passw)           
driver.find_element_by_xpath("/html/body/div[1]/form/div[3]/input[3]").click()

print('Apagar arquivos existentes')
for filename in os.listdir(downloadFilepath):
    if filename.startswith("AdesoesDetalhado_")|filename.startswith("Adesoes por periodo")|filename.startswith("CONTROLE_"):
        os.remove(os.path.join(filename))
    
driver.get(rel1)
print('Aguardar  a página carregar')
wait.until(EC.element_to_be_clickable((By.ID, 'ctl00_MainContent_RPV01_ctl04_ctl00')))
print('Preencher os dados e executar o relatório')
driver.find_element_by_xpath("/html/body/form/div[3]/div/div[2]/div/div[2]/div[4]/div/div/table/tbody/tr[2]/td/div/div/table/tbody/tr/td[1]/table/tbody/tr/td[2]/div/div/input[1]").send_keys (dt1)     
driver.find_element_by_xpath("/html/body/form/div[3]/div/div[2]/div/div[2]/div[4]/div/div/table/tbody/tr[2]/td/div/div/table/tbody/tr/td[1]/table/tbody/tr/td[5]/div/div/input[1]").send_keys (dt2)     
driver.find_element_by_xpath("/html/body/form/div[3]/div/div[2]/div/div[2]/div[4]/div/div/table/tbody/tr[2]/td/div/div/table/tbody/tr/td[3]/table/tbody/tr/td/input").click()                                

print('Esperar imagem de carregamento sumir')
wait.until(EC.invisibility_of_element_located((By.ID, "ctl00_MainContent_RPV01_AsyncWait_Wait")))
time.sleep(5)
print('Baixar arquivo de controle controle cartão')
driver.find_element_by_xpath("/html/body/form/div[3]/div/div[2]/div/div[2]/div[4]/div/div/table/tbody/tr[4]/td/div/div/div[2]/table/tbody/tr/td/div[1]/table/tbody/tr/td/a/img[2]").click()                                                     
driver.find_element_by_xpath("/html/body/form/div[3]/div/div[2]/div/div[2]/div[4]/div/div/table/tbody/tr[4]/td/div/div/div[2]/table/tbody/tr/td/div[2]/div[3]/a").click()                                                                    
            
print('Ir para a página do relatório de controle boleto')
driver.get(rel2)
print('Baixar arquivo do controle boleto')
driver.find_element_by_xpath("/html/body/form/div[3]/div/div[2]/div/div[1]/div[2]/ul/li[1]/a").click()

driver.close()

print('Verificar se o arquivo AdesoesDetalhado_* existe')
not_exist = True
while not_exist:
    for filename in os.listdir(downloadFilepath):
        if filename.startswith("AdesoesDetalhado_") and filename.endswith("csv"):
            print('Existe')
            not_exist = False
            break
        else: 
            print('Não existe')
            not_exist = True
    if not_exist:
        print('Aguarda 5 segundos')
        time.sleep(5)

print('Renomear arquivos')
for filename in os.listdir(downloadFilepath):
    if filename.startswith("AdesoesDetalhado_"):
        os.rename(os.path.join(filename), "CONTROLE_BOLETO")
    if filename.startswith("Adesoes por periodo"):
        os.rename(os.path.join(filename), "CONTROLE_CARTAO")