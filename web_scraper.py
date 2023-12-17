#importando selenium para extrair os dados, tkinter e PySimpleGUI para criar a interface gráfica, os - os.path, glob para trabalhar com os arquivos alem de time que é utilizado para criar um delay nas intruções que dependem do carregamento das páginas web
import tkinter as tk
from tkinter import ttk
import os
import glob
import os.path
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
import PySimpleGUI as sg

#Design da Interface
layout = [ 
    [sg.Radio("Exportação", "ie", key='expo')],
    [sg.Radio("Importação", "ie", key='impo')],
    [sg.Combo(['2018', '2019', '2020', '2021', '2022', '2023', '2024', '2025', '2026', '2027', '2028', '2029', '2030', '2031', '2032', '2033', '2034', '2035', '2036', '2037', '2038', '2039', '2040'],default_value='Ano?',key='board')],
    [sg.Radio("Financeiro", "documento", key='financeiro')],
    [sg.Radio("Operacional", "documento", key='operacional')],
    [sg.Radio("Todos", "documento", key='todos')],
    [sg.Text("Insira os processos como exemplificado abaixo:\nExemplo: 006/23,169/23,263/23\nSem espaço depois da virgula!")],
    [sg.InputText(key="input_processos")],
    [sg.Button("Baixar"), sg.Button("Cancelar")],
]

janela = sg.Window("Baixar Docs", layout)

#Loop para manter a janela até o usuario clicar em cancelar
while True:
    evento, valores = janela.read()
    if evento == sg.WIN_CLOSED or evento == "Cancelar":
        break
    if evento == "Baixar":
        fprocessos = valores["input_processos"]
        expo = valores["expo"]
        impo = valores["impo"]
        year = valores["board"]
        financeiro = valores["financeiro"]
        operacional = valores["operacional"]
        todos = valores["todos"]
        break

janela.close()

#Armazenando as opções escolhidas na interface em variaveis para uso posterior, são elas:
# - Importação ou Exportação
# - Tipo de documento desejado: Financeiro, operacional ou todos
if expo == True:
    ie = 'expo'
elif impo == True:
    ie = 'impo'

if financeiro == True:
    documento = 'Financeiro'
elif operacional == True:
    documento = 'Operacional'
elif todos == True:
    documento = 'Todos'

ano = year

processos = fprocessos.split(',')

LARGE_FONT= ("Verdana", 12)
NORM_FONT= ("Verdana", 10)
SMALL_FONT= ("Verdana", 8)

#Definindo o caminho das pastas 
if ie == 'impo':
    folders = os.listdir(f'U:\Arquivos Digitais - Comex\ARQUIVOS IMPO- ATUAL\ARQUIVOS IMPO- ATUAL\IMPORTAÇÕES {ano}')
elif ie == 'expo':
    folders = os.listdir(f'U:\\Arquivos Digitais - Comex\Arquivos digitais NAV EX\\{ano}\\')

#Configurações do webdriver
options = webdriver.ChromeOptions()
options.add_experimental_option("detach", True)
options.add_experimental_option("prefs", {
        "download.default_directory": "U:\\Comercio Exterior\\DOCUMENTOS SERVIMEX\\",
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing_for_trusted_sources_enabled": False,
        "safebrowsing.enabled": False
})
options.add_argument("--headless=new")

# Função para fazer o download dos processos desejados
def download(process, ie, ano, documento):
    driver = webdriver.Chrome(options=options)
    driver.get("https://www.servimex.com.br")
    driver.find_element(By.ID, 'submit').click() 
    driver.find_element(By.LINK_TEXT, "SERVIMEX ONLINE").click() #botao servimexonline da pagina inicial
    driver.implicitly_wait(1)
    driver.find_element(By.ID, "smLogin3").send_keys('#') #entraria com o login
    driver.find_element(By.ID, "smSenha3").send_keys('#') #entraria com a senha
    driver.find_element(By.ID, "envialogin").submit()
    time.sleep(2)
    frame = driver.find_elements(By.ID,'home')[0] 
    driver.switch_to.frame(frame)
    driver.find_element(By.XPATH, '/html/body/table/tbody/tr[3]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a/img').click() #clicar em consulta de processos
    if ie == 'impo':
        driver.find_element(By.NAME, 'f0').click()
        driver.find_element(By.XPATH, '/html/body/table/tbody/tr[3]/td/table/tbody/tr[1]/td/table/tbody/tr[2]/td/form/table/tbody/tr[1]/td[2]/select/option[2]').click()
    driver.find_element(By.NAME, 'f1').click() #display de opções - todos, em andamento e finalizado
    driver.find_element(By.XPATH, '/html/body/table/tbody/tr[3]/td/table/tbody/tr[1]/td/table/tbody/tr[2]/td/form/table/tbody/tr[1]/td[4]/select/option[3]').click() #muda situação para todos
    driver.find_element(By.NAME, "f3").send_keys(str(process)) # insere o processo
    driver.find_element(By.XPATH, '/html/body/table/tbody/tr[3]/td/table/tbody/tr[1]/td/table/tbody/tr[3]/td/table/tbody/tr/td[1]/a/img').click() #buscar
    if driver.find_element(By.XPATH, '/html/body/table/tbody/tr[3]/td/table/tbody/tr[1]/td/table/tbody/tr[3]/td/table/tbody/tr/td[2]/font/font').text == 'Não foi encontrado nenhum registro.':
        return popupmsg(f'o processo {process} não está disponivel na servimex')
    driver.find_element(By.XPATH, '/html/body/table/tbody/tr[3]/td/table/tbody/tr[2]/td/form[2]/table/tbody/tr/td/table/tbody/tr[2]/td[4]/a').click() #abre processo
    driver.find_element(By.LINK_TEXT, 'Documentos GED').click()
    if documento == 'Financeiro':
        driver.find_element(By.XPATH, '//*[@id="tabPane3"]/div[1]/h2[2]').click()
        iframe = driver.find_elements(By.NAME,'converge')[1]
    elif documento == 'Operacional':
        driver.find_element(By.XPATH, '//*[@id="tabPane3"]/div[1]/h2[1]').click()
        iframe = driver.find_elements(By.NAME,'converge')[0]
    elif documento == 'Todos':
        driver.find_element(By.XPATH, '//*[@id="tabPane3"]/div[1]/h2[1]').click()
        iframe = driver.find_elements(By.NAME,'converge')[0]
    driver.implicitly_wait(1)
    driver.switch_to.frame(iframe)
    driver.find_element(By.NAME,'ckb').click()
    driver.find_element(By.CLASS_NAME,'botoes').click()
    time.sleep(2)
    driver.quit()
    for folder in folders:
        if folder[:6]  == process[:6].replace("/", "-"):
            folder_path = r'U:\\Comercio Exterior\\DOCUMENTOS SERVIMEX\\'
            file_type = r'\*.zip'
            files = glob.glob(folder_path + file_type)
            max_file = max(files, key=os.path.getctime)
            old_name = "U:\\Comercio Exterior\\DOCUMENTOS SERVIMEX\\" + max_file[43:] #Mudar caminho dependendo.
            if ie == 'impo' and (documento == 'Operacional' or documento == 'Todos'):
                new_name = f"U:\Arquivos Digitais - Comex\ARQUIVOS IMPO- ATUAL\ARQUIVOS IMPO- ATUAL\IMPORTAÇÕES {ano}\\{folder}\\Operacional - {max_file[43:]}"
            elif ie == 'impo' and documento == 'Financeiro':
                new_name = f"U:\Arquivos Digitais - Comex\ARQUIVOS IMPO- ATUAL\ARQUIVOS IMPO- ATUAL\IMPORTAÇÕES {ano}\\{folder}\\Financeiro - {max_file[43:]}"
            elif ie == 'expo' and (documento == 'Operacional' or documento == 'Todos'):
                new_name = f"U:\\Arquivos Digitais - Comex\Arquivos digitais NAV EX\\{ano}\\{folder}\\Operacional - {max_file[43:]}" 
            elif ie == 'expo' and documento == 'Financeiro':
                new_name = f"U:\\Arquivos Digitais - Comex\Arquivos digitais NAV EX\\{ano}\\{folder}\\Financeiro - {max_file[43:]}"
            os.replace(old_name, new_name)
            break
    if documento == 'Todos':
        download(processo, ie, ano, 'Financeiro')

# Popup para sinalizar quando os processos inseridos forem finalizados
def popupmsg(msg):
    popup = tk.Tk()
    popup.wm_title("!")
    label = ttk.Label(popup, text=msg, font=NORM_FONT)
    label.pack(side="top", fill="x", pady=10)
    B1 = ttk.Button(popup, text="Okay", command = popup.destroy)
    B1.pack()
    popup.mainloop()

# Exibe o popup quando não há mais processos
for processo in processos:
    download(processo, ie, ano, documento)
    if processo == processos[-1]:
        popupmsg('Todos processos foram baixados e instalados!')
