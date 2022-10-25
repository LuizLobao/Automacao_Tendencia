from cgitb import text
import shutil,os,time
from tkinter import *
from tkinter import ttk
from datetime import date, datetime
from PySimpleGUI import Window, Button, Text, Image, Column, VSeparator, Push, theme, popup, MenuBar
from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep

hoje = datetime.today().strftime('%Y-%m-%d %H:%M:%S')
#Puxa o Login do usuário 
login = os.getlogin()
#Nome do Arquivo Demonstrativo Gross
arquivo1 = 'Demonstrativo Gross'


def puxa_conteudo_online():
    driver = webdriver.Edge(r'C:\\Users\\oi066724\\Documents\\Python\\Python\\msedgedriver.exe')
    #Endereço do Monitor de Cargas
    driver.get('http://10.20.83.116/aplicacao/monitor/')
    sleep(3)
    val_1066 = driver.find_element(By.XPATH,"/html/body/table[2]/tbody/tr[7]/td[8]").text
    val_1067 = driver.find_element(By.XPATH,"/html/body/table[2]/tbody/tr[8]/td[8]").text
    val_1059 = driver.find_element(By.XPATH,"/html/body/table[2]/tbody/tr[9]/td[8]").text
    val_1064 = driver.find_element(By.XPATH,"/html/body/table[2]/tbody/tr[10]/td[8]").text
    val_1065 = driver.find_element(By.XPATH,"/html/body/table[2]/tbody/tr[11]/td[8]").text
    val_1058 = driver.find_element(By.XPATH,"/html/body/table[2]/tbody/tr[12]/td[8]").text
    val_6162 = driver.find_element(By.XPATH, "/html/body/table[2]/tbody/tr[82]/td[8]").text
    val_6163 = driver.find_element(By.XPATH, "/html/body/table[2]/tbody/tr[82]/td[8]").text


#puxa_conteudo_online()

# Verificar data de Modificação antes de Realizar a Cópia do Arquivo
def data_mod_arquivo():
    arquivo = f'Y:\{arquivo1}.xlsb'
    #modificado = ("last modified: %s" % time.ctime(os.path.getmtime(arquivo)))
    modificado = (time.ctime(os.path.getmtime(arquivo)))
    #print(modificado)
    #print("created: %s" % time.ctime(os.path.getctime(arquivo)))
    return (modificado)

# Realiza a copia do Arquivo para a pasta de Donwload local renomeando o arquivo
def copia_arquivo_renomeia():
    data = datetime.today().strftime('%Y%m%d')
    #shutil.copy("Y:\Demonstrativo Gross.xlsb", 'C:\Users\oi066724\Downloads\Demonstrativo Gross_teste.xlsb')


############################## MONTANDO A JANELA DE INTERFACE COM O USUÁRIO ##############################
theme('DarkBlue')

# ------ Menu Definition ------ #
menu_def = [
    ['&Arquivo', ['&Abrir', '&Salvar', '---', 'Propriedades', '&Sair'  ]],
    ['&Editar', ['Colar', ['Especial', 'Normal',], 'Desfazer'],],
    ['A&juda', '&Sobre...']
]

layout_esquerda = [
   # [Image(filename='fusion.png')]
]
layout_direita = [
    [Text(f'Olá {login}, seja bem vindo:')],                                                                                #Linha 0
    [Text()],                                                                                                               #Linha em branco
    [Text(f'Data de modificação do arquivo {arquivo1} :'), Text(f'{data_mod_arquivo()}:'), Button('Copiar Arquivo',key='-COPIAR-')],
    [Text()],                                                                                                               #Linha em branco
    [Text('Puxar arquivo da Intranet'), Button('puxar arquivo', key = '-PUXARARQUIVO-')],                                   #Linha em branco
    [Text()],                                                                                                               #Linha em branco
    [Button('btn1'), Push(), Button('btn2'), Push(), Button('btn3')]                                                        #Linha 2
]
layout = [
    [MenuBar(menu_def), Column(layout_esquerda),VSeparator(), Column(layout_direita)]
]



window = Window(
    f'Sistema de Automação de Tendências - {hoje}',
    #size=(500,300),
    layout=layout
)
#print(window.read())

while True:
    event,values = window.read()

    match(event):
        case '-COPIAR-':
            copia_arquivo_renomeia(),
            popup('Arquivo Copiado')
        case '-PUXARAQUIVO-':
            puxa_conteudo_online()
        case None:
            break

window.close()

