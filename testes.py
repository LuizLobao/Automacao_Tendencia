import segredos
import pyodbc
from playwright.sync_api import sync_playwright
from datetime import date, datetime
import time
import pandas as pd
import win32com.client as win32
import plotly.express as px

hoje = datetime.today().strftime('%d/%m/%Y')
print(hoje)



def testepandas():
    comando_sql='''
				SELECT * FROM TBL_RE_BASERESULTADOS 
                WHERE DATA = '202210' and 
                tipo_indicador in ('real','tendência') and 
                indbd = 'VL' and
                grupo_plano in ('Fibra', 'Nova Fibra')
                '''

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
    df=pd.read_sql(comando_sql, conexao)
    #print(df.head())
    #print(df.tail())
    pt_tabdin = df.pivot_table(
                                    values="VALOR", 
                                    index=["FILIAL"], 
                                    columns="TIPO_INDICADOR", 
                                    aggfunc=sum,
                                    fill_value=0,
                                    margins=True, margins_name="VL",
                                    )
    conexao.close()
    print(pt_tabdin)
    fig = px.bar(df, x='DATA', y='VALOR')
    fig.show()
    print('Conexão Fechada')

testepandas()