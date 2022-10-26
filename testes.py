import segredos
import pyodbc
from playwright.sync_api import sync_playwright
from datetime import date, datetime
import pandas as pd
import win32com.client as win32
import plotly.express as px

hoje = datetime.today().strftime('%d/%m/%Y')
agora = datetime.today()

dados_conexao = (
        "Driver={SQL Server};"
        f"Server={segredos.db_server};"
        f"Database={segredos.db_name};"
        f"UID={segredos.db_user};"
        f"PWD={segredos.db_pass}"
    )
conexao = pyodbc.connect(dados_conexao)
print("Conectado")

cursor = conexao.cursor()

#colocar aqui um loop verificando se a procedure terminou de rodar
# gerar variavel com a datahora de inicio de execuçao inicio = datetime.today()
# a procedure ira gravar a hora de inicio e fim da execução na tabela TBL_PC_TEMPO_PROCEDURES
# puxar a hora FIM da procedure
# criar um loop que ira verificar se a hora de fim da procedure for MAIOR que a variavel INICIO
# se MENOR, continuar esperando
# se MAIOR, continuar com o processo

print('Iniciando LOOP para verificar se a Procedure terminou de rodar')
comando_sql = '''SELECT [PROCEDURE], INI_FIM, MAX(DATA_HORA) AS DATA_HORA
				  FROM TBL_PC_TEMPO_PROCEDURES
   			  WHERE [PROCEDURE] = 'SP_PC_BASES_SHAREPOINT'AND INI_FIM = 'FIM'
				  GROUP BY [PROCEDURE], INI_FIM'''


cursor.execute(comando_sql)
row = cursor.fetchone()
print(row)
print(row.DATA_HORA)
print(row.DATA_HORA.strftime('%Y%m%d'))
print(agora)

print(row.DATA_HORA>agora)

conexao.close()
print('Conexão Fechada')