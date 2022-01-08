import os, time
from pathlib import Path

from openpyxl import Workbook
from openpyxl import load_workbook

from PIL import ImageGrab
import win32com.client as win32

from selenium import webdriver
from selenium.webdriver import Firefox
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.service import Service

from datetime import datetime

import logging

start_time = datetime.now()

logging.basicConfig(filename=str(Path(__file__).parent.absolute()) + "\\" + "historico.log", level=logging.INFO, format='%(message)s')
logging.info("")
logging.info("==================== INICIO ====================")
logging.info("")
logging.info("Começo do Script: " + str(datetime.now().strftime(r'%Y-%m-%d %H:%M:%S')))
logging.info("")

#=== CRIANDO UM PROFILE NO FIREFOX ===#

# Windows + R
# "firefox.exe -p"
# Crie um novo perfil
# Entre na pasta: "%appdata%\Roaming\Mozilla\Firefox\Profiles"
# Copie o caminho para a variavel "profile_path"

lista_contatos = []
lista_figuras = []

    #=== CONFIGURACOES ===#

#=== CAMINHO DO WEBDRIVER E DA PASTA PROFILE DO FIREFOX ===#
driver_path = str(Path(__file__).parent.absolute()) + r"\geckodriver.exe" #caminho do navegador
profile_path = r"C:\Users\higor\AppData\Roaming\Mozilla\Firefox\Profiles\<seu_profile>" #caminho do profile, geralmente localizado em "%appdata%\Roaming\Mozilla\Firefox\Profiles" *fazer login no whatsapp antes

#=== NOME DA PLANILHA USADA PARA RODAR O ENVIO ===#
planilha_contatos = str(Path(__file__).parent.absolute()) + r"\Exemplo.xlsx" #nessa planilha os contatos ficam a esquerda e os nomes das imagens ficam a direita, tudo na mesma planilha
aba_planilha = "Aba1"

#=== VARIAVEIS P/ DEFINIR ONDE COMEÇAM AS LISTAS DE CONTATOS E FIGURAS ===#

coluna_contatos = 1 #coluna A
linha_contatos = 2 #linha 2
coluna_figuras = 2 #coluna B
linha_figuras = 2 #linha 2

    #=== INICIANDO O NAVEGADOR ===#
options = Options()
options.headless = False #true pra rodar o firefox em modo oculto
options.add_argument("-profile") 
options.add_argument(profile_path) #adicionando a pasta profile p/ o whatsapp não ficar pedindo o QR code  

service = Service(driver_path) #caminho do driver

driver = Firefox(service=service, options=options) 

driver.get("http://web.whatsapp.com")

wait = WebDriverWait(driver, 800)

def importar_contatos(planilha):
    wb = load_workbook(planilha)

    try:
        ws = wb[aba_planilha] #nome da aba da planilha
        
        for row in range(linha_contatos, ws.max_row+1): #numero da linha onde começa / maximo de linhas +1 para não finalizar antes do ultimo elemento   
            if(ws.cell(row, coluna_contatos).value is None): #se valor é nulo -> break
                break
            else:
                lista_contatos.append(ws.cell(row, coluna_contatos).value) #concatena o valor encontrado na lista
    except Exception as e:
        print("Erro no modulo 'importar_contatos': " + str(e))
        logging.error(str(datetime.now().strftime(r'%H:%M:%S')) + " === Erro no modulo 'importar_contatos': " + str(e))
    finally:
        wb.close() #fecha o workbook p/ win32com conseguir copiar a imagem

def importar_figura(planilha):
    wb = load_workbook(planilha)

    try:
        ws = wb[aba_planilha] #nome da aba da planilha
    
        for row in range(linha_figuras, ws.max_row+1): #numero da linha onde começa / maximo de linhas +1 para não finalizar antes do ultimo elemento
            if(ws.cell(row, coluna_figuras).value is None): #se valor é nulo -> break
                break
            else:
                lista_figuras.append(ws.cell(row, coluna_figuras).value) #define na variavel o nome da imagem
    except Exception as e:
        print("Erro no modulo 'importar_figura': " + str(e))
        logging.error(str(datetime.now().strftime(r'%H:%M:%S')) + " === Erro no modulo 'importar_figura': " + str(e))
    finally:
        wb.close() #fecha o workbook p/ win32com conseguir copiar a imagem

def salvar_imagem(planilha):
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    wb = excel.Workbooks.Open(planilha)

    try:
        sheet = excel.Sheets(aba_planilha)

        for i, shape in enumerate(sheet.Shapes): #percorre todos os shapes da pasta selecionada
            for figura in lista_figuras: #percorre todas as figuras armazenadas na lista de figuras
                if shape.Name.startswith(figura): #se o nome da figura for igual ao elemento da lista de figuras
                    shape.Copy()
                    image = ImageGrab.grabclipboard() #copia
                    image = image.convert('RGB') #converte de RGBA p/ RGB
                    image.save(str(Path(__file__).parent.absolute()) + "\\" + figura + ".jpg", "jpeg") #salva com a extensao .jpg
    except Exception as e:
        print("Erro no modulo 'salvar_imagem': " + str(e))
        logging.error(str(datetime.now().strftime(r'%H:%M:%S')) + " === Erro no modulo 'salvar_imagem': " + str(e))
    finally:
        wb.Close(True) #fecha o workbook
        excel.Quit() #finaliza o processo do excel

def envia_imagens(contatos, imagens):
    print("Enviando mensagens...")
    logging.info(str(datetime.now().strftime(r'%H:%M:%S')) + " === Enviando mensagens...")
    
    time.sleep(5)
    try:
        for i in range(0, len(contatos)): #percorre do 0 até o tamanho do array dos contatos enviando para o contato a respectiva imagem na mesma linha da planilha
            print("Enviando imagem: ## " + str(imagens[i].upper()) + " ## Para: ## " + str(contatos[i]).upper() + " ##")
            logging.info(str(datetime.now().strftime(r'%H:%M:%S')) + " === Enviando imagem: ## " + str(imagens[i].upper()) + " ## Para: ## " + str(contatos[i]).upper() + " ##")
            #procura na pagina o elemento com o mesmo nome do contato
            x_arg = '//span[contains(@title, ' + '"' + str(contatos[i]) + '"' + ')]'
            group_title = wait.until(EC.presence_of_element_located((
                By.XPATH, x_arg)))
            group_title.click()
            #procura na pagina o elemento de anexo de arquivos
            driver.find_element(By.CSS_SELECTOR, "span[data-icon='clip']").click()
            attach = driver.find_element(By.CSS_SELECTOR, "input[type='file']")
            #passa o caminho da imagem para ser anexada
            attach.send_keys(str(Path(__file__).parent.absolute()) + "\\" + str(imagens[i]) + ".jpg")
            time.sleep(3)
            #tecla enviar
            send = driver.find_element(By.CSS_SELECTOR, "span[data-icon='send']")
            send.click()
            time.sleep(5)
    except Exception as e:
        print("Erro no modulo 'envia_imagens': " + str(e))
        logging.error(str(datetime.now().strftime(r'%H:%M:%S')) + " === Erro no modulo 'envia_imagens': " + str(e))
    finally:
        time.sleep(3)
        driver.close() #fecha o driver

def rotina():
    #=== EXECUTANDO FUNÇÕES ===#
    #importando contatos
    importar_contatos(planilha_contatos) #passando a planilha carregada antes como argumento
    print("Lista de contatos importada:") 
    print(lista_contatos)
    logging.info(str(datetime.now().strftime(r'%H:%M:%S')) + " === Lista de contatos importada:") 
    logging.info(lista_contatos)

    #importando figuras
    importar_figura(planilha_contatos) #passando a planilha carregada antes como argumento
    print("Lista de figuras importadas:")
    print(lista_figuras)
    logging.info(str(datetime.now().strftime(r'%H:%M:%S')) + " === Lista de figuras importadas:")
    logging.info(lista_figuras)

    salvar_imagem(planilha_contatos) #passando a planilha carregada antes como argumento

    envia_imagens(lista_contatos, lista_figuras) #passa a lista de contatos e a lista de figuras armazenadas nos arrays como argumentos

    #excluindo as figuras geradas
    for figura in lista_figuras:
        logging.info(str(datetime.now().strftime(r'%H:%M:%S')) + " === Deletando arquivo: " + str(Path(__file__).parent.absolute()) + "\\" + str(figura) + ".jpg")
        os.remove(str(Path(__file__).parent.absolute()) + "\\" + str(figura) + ".jpg")  

#=== VERIFICANDO SE EXISTE O ARQUIVO EXCEL ===#

if os.path.exists(planilha_contatos):
    rotina()
else:
    print("Arquivo Excel nao existe!")
    logging.info(str(datetime.now().strftime(r'%H:%M:%S')) + " === Arquivo Excel nao existe! Verifique se '" + planilha_contatos + "' esta correto.")

end_time = datetime.now()

logging.info("")
logging.info("Final do Script: " + str(datetime.now().strftime(r'%Y-%m-%d %H:%M:%S')))
logging.info("Tempo de execução: {}".format(end_time - start_time))
print('---- Tempo de execução: {} ----'.format(end_time - start_time))

logging.info("")
logging.info("====================== FIM ======================")
logging.info("")

print(str(Path(__file__).parent.absolute()))