import segredos
from datetime import date, datetime
import pyodbc
import time


def executa_procedure_sql_1():
   
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
    
    #executa procedure
    inicio_procedure = datetime.today()
    parametros = datetime.today().strftime('%Y%m')
    print(parametros)
    print('\x1b[1;33;44m' + 'Executando a Procedure SP_PC_TEND_IGUAL_REAL_FIBRA_EMPRESARIAL'+ '\x1b[0m')
    cursor.execute(f'SET NOCOUNT ON; EXEC SP_PC_TEND_IGUAL_REAL_FIBRA_EMPRESARIAL {parametros}')
    fim_procedure = datetime.today()
    conexao.commit()

    print(f"Procedure executada em {fim_procedure - inicio_procedure} tempo")
    
    conexao.close()
    print('Conexão Fechada')

def executa_procedure_sql_2():
      
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
    
    #executa procedure
    inicio_procedure = datetime.today()
    parametros = datetime.today().strftime('%Y%m')
    print(parametros)
    print('\x1b[1;33;44m' + 'Executando a Procedure SP_PC_TEND_IGUAL_REAL_FIBRA_VAREJO'+ '\x1b[0m')
    cursor.execute(f'SET NOCOUNT ON; EXEC SP_PC_TEND_IGUAL_REAL_FIBRA_VAREJO {parametros}')
    fim_procedure = datetime.today()
    conexao.commit()

    print(f"Procedure executada em {fim_procedure - inicio_procedure} tempo")
    
    conexao.close()
    print('Conexão Fechada')

def executa_procedure_sql_3():
      
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
    
    #executa procedure
    inicio_procedure = datetime.today()
    parametros = datetime.today().strftime('%Y%m')
    print(parametros)
    print('\x1b[1;33;44m' + 'Executando a Procedure SP_PC_TEND_IGUAL_REAL_NOVA_FIBRA'+ '\x1b[0m')
    cursor.execute(f'SET NOCOUNT ON; EXEC SP_PC_TEND_IGUAL_REAL_NOVA_FIBRA {parametros}')
    fim_procedure = datetime.today()
    conexao.commit()

    print(f"Procedure executada em {fim_procedure - inicio_procedure} tempo")
    
    conexao.close()
    print('Conexão Fechada')

def executa_procedure_sql_4():
      
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
    
    #executa procedure
    inicio_procedure = datetime.today()
    parametros = datetime.today().strftime('%Y%m')
    print(parametros)
    print('\x1b[1;33;44m' + 'Executando a Procedure SP_PC_TEND_IGUAL_REAL_TABELAS_FIBRA'+ '\x1b[0m')
    cursor.execute(f'SET NOCOUNT ON; EXEC SP_PC_TEND_IGUAL_REAL_TABELAS_FIBRA {parametros}')
    fim_procedure = datetime.today()
    conexao.commit()

    print(f"Procedure executada em {fim_procedure - inicio_procedure} tempo")
    
    conexao.close()
    print('Conexão Fechada')    


executa_procedure_sql_1()
executa_procedure_sql_2()
executa_procedure_sql_3()
executa_procedure_sql_4()