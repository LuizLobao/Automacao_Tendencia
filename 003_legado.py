# 1) Verifica se as bases do Legado já carregaram
# 2) Executar procedure exec SP_PC_Insert_Tendencia_Auto_Fibra 202210 --passar o AAAAMM como parametro na procedure
# 3) Abrir o Excel para continuar os ajustes manuais
# 4) Rodar consultas para igualar REAL e TENDÊNCIA

import segredos
from playwright.sync_api import sync_playwright
from datetime import date, datetime
import pyodbc
import time

hoje = datetime.today().strftime('%d/%m/%Y')

def puxa_dts_cargas():
    with sync_playwright() as p:
        navegador = p.chromium.launch(headless=True)
        pagina = navegador.new_page()
        pagina.goto("http://10.20.83.116/aplicacao/monitor/")
        
        datafim = {pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[10]/td[9]').text_content() : pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[10]/td[8]').text_content(),
                   pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[7]/td[9]').text_content() : pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[7]/td[8]').text_content(),
                   pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[9]/td[9]').text_content() : pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[9]/td[8]').text_content(),
                   pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[8]/td[9]').text_content() : pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[8]/td[8]').text_content(),
                   pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[12]/td[9]').text_content() : pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[12]/td[8]').text_content(),
                   pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[11]/td[9]').text_content() :pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[11]/td[8]').text_content()   
                   }
        navegador.close()
        return datafim

def executa_procedure_sql(nome_procedure, param):
   
    dados_conexao = (
        "Driver={SQL Server};"
        f"Server={segredos.db_server};"
        f"Database={segredos.db_name};"
        f"UID={segredos.db_user};"
        f"PWD={segredos.db_pass}"
    )
    conexao = pyodbc.connect(dados_conexao)
    print('\x1b[1;33;42m' + 'Conexão realizada ao banco de dados' + '\x1b[0m')

    cursor = conexao.cursor()
    
    #executa procedure
    inicio_procedure = datetime.today()
    print('\x1b[1;33;44m' + f'Executando a Procedure {nome_procedure} para o parâmetro: {param} '+ '\x1b[0m')
    print(f'Iniciando execução em: {inicio_procedure}')
    cursor.execute(f'SET NOCOUNT ON; EXEC {nome_procedure}  {param}')
    conexao.commit()
    fim_procedure = datetime.today()
    print(f"Procedure executada em {fim_procedure - inicio_procedure} tempo")
    
    conexao.close()
    print('\x1b[1;33;41m' + 'Conexão Fechada'+ '\x1b[0m')


fim = puxa_dts_cargas()

BOV_1058 = (f'{fim["BOV_1058.TXT"].split(" ")[0]}')
BOV_1059 = (f'{fim["BOV_1059.TXT"].split(" ")[0]}')
BOV_1064 = (f'{fim["BOV_1064.TXT"].split(" ")[0]}')
BOV_1065 = (f'{fim["BOV_1065.TXT"].split(" ")[0]}')
BOV_1066 = (f'{fim["BOV_1066.TXT"].split(" ")[0]}')
BOV_1067 = (f'{fim["BOV_1067.TXT"].split(" ")[0]}')


print(f'HOJE    : {hoje}')
print(f'BOV_1058: {BOV_1058}')
print(f'BOV_1059: {BOV_1059}')
print(f'BOV_1064: {BOV_1064}')
print(f'BOV_1065: {BOV_1065}')
print(f'BOV_1066: {BOV_1066}')
print(f'BOV_1067: {BOV_1067}')

if hoje == BOV_1058 == BOV_1059 == BOV_1064 == BOV_1065 == BOV_1066 == BOV_1067:
    print('CONTINUANDO')
    param = datetime.today().strftime('%Y%m')
    executa_procedure_sql('SP_PC_Insert_Tendencia_Auto_Fibra',param)
    
