import segredos
from datetime import date, datetime
import pyodbc
import time


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


#param = datetime.today().strftime('%Y%m')
param = '202211'
executa_procedure_sql('SP_PC_TEND_IGUAL_REAL_FIBRA_EMPRESARIAL',param)
executa_procedure_sql('SP_PC_TEND_IGUAL_REAL_FIBRA_VAREJO',param)
executa_procedure_sql('SP_PC_TEND_IGUAL_REAL_NOVA_FIBRA',param)
executa_procedure_sql('SP_PC_TEND_IGUAL_REAL_TABELAS_FIBRA',param)