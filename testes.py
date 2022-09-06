#https://www.youtube.com/watch?v=AL8gqZ1jJxs

import pyodbc

server = 'SQLPW90DB03\DBINST3, 1443' 
database = 'BDintelicanais' 
username = 'INFOCANA' 
password = 'inf@C4N4IS' 



dados_conexao = (
    "DRIVER={SQL Server};"
    "SERVER=SQLPW90DB03\DBINST3, 1443;"
    "DATABASE=BDintelicanais;"
    "UID=INFOCANA;"
    "PWD=inf@C4N4IS;"
)

conexao = pyodbc.connect(dados_conexao)
print ("Conexao OK")
cursor = conexao.cursor()

comando = """
            select top 100 * from tbl_re_baseresultados
"""

cursor.execute(comando)