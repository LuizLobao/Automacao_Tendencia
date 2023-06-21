import shutil,os,time,pyodbc, segredos, openpyxl, subprocess, requests
import win32com.client as win32
import pandas as pd
from urllib.parse import quote
#from telnetlib import theNULL
from datetime import date, datetime, timedelta
from playwright.sync_api import sync_playwright
from tqdm import tqdm
from PIL import ImageGrab



def verifica_max_data_cdo_real():
	dados_conexao = (
		"Driver={SQL Server};"
		f"Server={segredos.db_server};"
		f"Database={segredos.db_name};"
		f"UID={segredos.db_user};"
		f"PWD={segredos.db_pass}"
	)
	conexao = pyodbc.connect(dados_conexao)
	print("Conectado ao banco para alterar a procedure - retirar comentários")
	
        # Defina o caminho do arquivo .sql
	caminho_arquivo_sql = r'S:\\01-Projetos\\03-CDO\\verifica_max_data_real.sql'
        
	with open(caminho_arquivo_sql, 'r', encoding='utf-8') as arquivo:
		conteudo_sql = arquivo.read()
		cursor = conexao.cursor()
		cursor.execute(conteudo_sql)
		resultado = cursor
		columns = [column[0] for column in cursor.description]
		#print(columns)
		results = []
		for row in cursor.fetchall():
			results.append(dict(zip(columns,row)))
		#print(results)
		#print(type(resultado))
	conexao.commit()
	conexao.close()
	print('Conexão Fechada')
	return(results)
	
a = verifica_max_data_cdo_real()
for itens in a:
    print(f'{itens["DS_PRODUTO"]} - {itens["DS_INDICADOR"]} = {itens["YYYYMMDD"]}')
