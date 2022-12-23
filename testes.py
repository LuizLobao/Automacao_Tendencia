import segredos
import pyodbc
from playwright.sync_api import sync_playwright
from datetime import date, datetime
import pandas as pd
import win32com.client as win32
import plotly.express as px
import time


data = pd.read_csv(r'C:\Users\oi066724\Downloads\BOVs\unzip\HADOOP_6163_488157_RDA_CLIENTCO_GROSS20221201_20221218_213509.TXT', sep="\t")
df = pd.DataFrame(data)

print(df)

#dados_conexao = (
#    "Driver={SQL Server};"
#    f"Server={segredos.db_server};"
#    f"Database={segredos.db_name};"
#    f"UID={segredos.db_user};"
#    f"PWD={segredos.db_pass}"
#)
#conexao = pyodbc.connect(dados_conexao)
#print("Conectado ao banco para executar PROCEDURE")
#
#cursor = conexao.cursor()