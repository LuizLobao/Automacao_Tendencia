#antes de começar
#python -m pip install statsmodels

import itertools
import pandas as pd
import pyodbc
import segredos
import statsmodels.api as sm

from statsmodels.tsa.holtwinters import ExponentialSmoothing
from statsmodels.tsa.seasonal import seasonal_decompose



def criar_conexao():
    dados_conexao = (
        "Driver={SQL Server};"
        f"Server={segredos.db_server};"
        f"Database={segredos.db_name};"
        "Trusted_Connection=yes;"
        # f"UID={segredos.db_user};"
        # f"PWD={segredos.db_pass}"
    )
    return pyodbc.connect(dados_conexao)

def monta_sql (indicador, gestao, produto, un):
    cod_sql=f'''
        SELECT CONVERT(DATE,[DT_REFERENCIA],103) AS DATA
        ,SUM(CAST([QTD] AS FLOAT)) AS QTD
        FROM [BDINTELICANAIS].[DBO].[TBL_CDO_FISICOS_REAL] AS A
        LEFT JOIN DBO.TBL_CDO_DE_PARA_REGIONAL AS B ON A.NO_CURTO_TERRITORIO = B.NO_CURTO_TERRITORIO
        LEFT JOIN DBO.TBL_CDO_DE_PARA_CANAL AS C ON A.DS_CANAL_BOV = C.DS_DESCRICAO_CANAL_BOV
        WHERE [DS_DET_INDICADOR] IN ('NOVOS CLIENTES','MIG AQUISICAO')
        AND (CASE WHEN C.DS_GESTAO = 'GESTAO REGIONAL' THEN B.COD_REGIONAL
                WHEN C.DS_GESTAO = 'OUTROS NACIONAIS' THEN 'OUTROS'
                WHEN (C.DS_GESTAO = 'GESTAO NACIONAL' AND C.DS_CANAL_FINAL LIKE '%TLV%') THEN 'TLV'
                WHEN (C.DS_GESTAO = 'GESTAO NACIONAL' AND C.DS_CANAL_FINAL LIKE '%WEB%') THEN 'WEB'
                ELSE 'NAO CLASSIFICADO'
                END) = '{gestao}'
        AND A.DS_PRODUTO = '{produto}'
        AND A.DS_INDICADOR = '{indicador}'
        AND A.DS_UNIDADE_NEGOCIO = '{un}'
        GROUP BY CONVERT(DATE,[DT_REFERENCIA],103)
        ORDER BY 1
    '''
    return(cod_sql)



indicador = ['VL','VLL','GROSS']
gestao = ['TLV','WEB','RCS','RSE','RNN','OUTROS']
produto = ['FIBRA','NOVA FIBRA']
un = ['VAREJO','EMPRESARIAL']

comb_list = list(itertools.product(indicador, gestao, produto, un))
#print(comb_list)

#for itens in comb_list:
#    print(itens[0])
#    print(itens[1])
#    print(itens[2])

comando_sql = monta_sql('VL', 'RNN', 'FIBRA', 'EMPRESARIAL')
#print(comando_sql)

conexao = criar_conexao()
df=pd.read_sql(comando_sql, conexao, parse_dates =['DATA'])
conexao.close()
df=df.groupby('DATA').sum()
df=df.resample(rule='D').sum()
print(df.head())
print(df.tail())


#train and test
#train = df[:516] #achar uma solução para pegar sempre até o ultimo dia do mÊs anterior
#test = df[516:]

#final_model=ExponentialSmoothing(df.QTD, trend='mul', seasonal='add', seasonal_periods=7).fit()
#pred=final_model.forecast(27) #mudar a quantidade para gerar até o final do mÊs corrente

#print(pred)