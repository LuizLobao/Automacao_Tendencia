import segredos
import pyodbc
from datetime import date, datetime
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


param = datetime.today().strftime('%Y%m')



proc = 'SP_PC_Update_Ticket_Fibra_VAREJO_Tendencia_porRegiao'
executa_procedure_sql(proc, param)


proc = 'SP_PC_Update_Ticket_Fibra_EMPRESARIAL_Tendencia_porRegiao_IndCombo'
executa_procedure_sql(proc, param)


proc = 'SP_PC_TBL_RE_RELATORIO_RC_V2_TEND'
executa_procedure_sql(proc, param)