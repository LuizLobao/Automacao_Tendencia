# 1) Verificar se bases rodaram no MONITOR DE CARGA - OK
# 2) Rodar procedure com codigo "descomentado" - OK
# 3) Rodar procedure SP_PC_NOVA_FIBRA - OK
# 4) Rodar procedure com codigo comentado - OK
# 5) Rodar query e colocar no excel - OK
# 6) salvar na rede - OK
# 7) mandar por e-mail - OK

import segredos
import pyodbc
from playwright.sync_api import sync_playwright
from datetime import date, datetime
import time
import pandas as pd
import win32com.client as win32

hoje = datetime.today().strftime('%d/%m/%Y')
AAAAMMDD = datetime.today().strftime('%Y%m%d')
print(hoje)

def puxa_dts_cargas():
    with sync_playwright() as p:
        navegador = p.chromium.launch(headless=True)
        pagina = navegador.new_page()
        pagina.goto("http://10.20.83.116/aplicacao/monitor/")
        
        dt6163 = pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[82]/td[8]').text_content()
        dt6162 = pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[83]/td[8]').text_content()

        dw6163 = pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[82]/td[5]').text_content()
        dw6162 = pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[83]/td[5]').text_content()

        di6163 = pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[82]/td[7]').text_content()
        di6162 = pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[83]/td[7]').text_content()

        navegador.close()

        datas_fim = [dt6163,dt6162]
        datas_down = [dw6163,dw6162]
        datas_ini = [di6163,di6162]

        #print(dt6163)
        
        return datas_fim, datas_down, datas_ini

def executa_procedure_sql():
    #Procedure para rodar 'exec SP_PC_NOVA_FIBRA'
        
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
    cursor.execute("{CALL SP_PC_NOVA_FIBRA}")
    conexao.commit()

    #colocar aqui um loop verificando se a procedure terminou de rodar
	#print(f"Iniciando LOOP para verificar se a Procedure terminou de rodar")
	#comando_sql = '''SELECT [PROCEDURE], INI_FIM, MAX(DATA_HORA) AS DATA_HORA
	#				  FROM TBL_PC_TEMPO_PROCEDURES
	#   			  WHERE [PROCEDURE] = 'SP_PC_BASES_SHAREPOINT'AND INI_FIM = 'FIM'
	#				  GROUP BY [PROCEDURE], INI_FIM'''
	#df=pd.read_sql(comando_sql, conexao)
    


    conexao.close()
    print('Conexão Fechada')

def montaExcelTendVll():
    comando_sql = '''SELECT DATA,
                    FILIAL,
                    SUM(QTD) AS [TEND]
                    FROM tbl_pc_tend_nova_fibra
                    GROUP BY DATA, FILIAL'''
    dados_conexao = (
        "Driver={SQL Server};"
        f"Server={segredos.db_server};"
        f"Database={segredos.db_name};"
        f"UID={segredos.db_user};"
        f"PWD={segredos.db_pass}"
    )
    conexao = pyodbc.connect(dados_conexao)
    #print("Conectado")
    cursor = conexao.cursor()
    df=pd.read_sql(comando_sql, conexao)
    pt_tabdin = df.pivot_table(
                                    values="TEND", 
                                    index=["FILIAL"], 
                                    columns="DATA", 
                                    aggfunc=sum,
                                    fill_value=0,
                                    margins=True, margins_name="VLL",
                                    )
    dest_filename = (f'S:\\Resultados\\01_Relatorio Diario\\1 - Base Eventos\\02 - TENDÊNCIA\\Insumos_Tendência\\Tend_VLL_Nova_Fibra_{AAAAMMDD}.xlsx')
    with pd.ExcelWriter(dest_filename) as writer:
            df.to_excel(writer, sheet_name="NOVA FIBRA VLL DADOS",startcol=0, startrow=0, index=0)
            pt_tabdin.to_excel(writer, sheet_name="PROJ. NOVA FIBRA VLL",startcol=0, startrow=0, index=1)

def enviaEmaileAnexo():        
    # criar a integração com o outlook
    outlook = win32.Dispatch('outlook.application')

    # criar um email
    email = outlook.CreateItem(0)

    # configurar as informações do seu e-mail
    email.To = segredos.lista_email_vll_nf_to
    email.Cc = segredos.lista_email_vll_nf_cc
    email.Subject = f"Projeção NOVA FIBRA - {hoje}"
    email.HTMLBody = f"""
    <p>Caros,</p>

    <p>Segue o arquivo atualizado com a projeção de VLL da Nova Fibra calculada hoje: {hoje}</p>
    <p></p>
    <p></p>

    <p>Att,</p>
    <p>Luiz Lobão</p>
    """
    anexo = (f'S:\\Resultados\\01_Relatorio Diario\\1 - Base Eventos\\02 - TENDÊNCIA\\Insumos_Tendência\\Tend_VLL_Nova_Fibra_{AAAAMMDD}.xlsx')
    email.Attachments.Add(anexo)

    email.Send()
    print("Email Enviado")

def tira_comentario_procedure_nova_fibra_sql():
    comando_sql='''
				ALTER PROCEDURE [dbo].[SP_PC_NOVA_FIBRA] AS

				DECLARE @ANOMES AS VARCHAR(6)
				SET @ANOMES = (SELECT DISTINCT DBO.FORMAT_DATE(DT_MES,'YYYYMM') FROM TBL_CG_NOVA_FIBRA_VL)

				-----------------------------------------------------------------------------------------
				-------------------- REGISTRO DE INICIO DE EXECUÇÃO -----------------------------
				-----------------------------------------------------------------------------------------
				INSERT INTO TBL_PC_TEMPO_PROCEDURES
				SELECT
				'SP_PC_NOVA_FIBRA' AS [PROCEDURE],
				'INICIO' AS INI_FIM,
				GETDATE() AS DATA_HORA

				-----------------------------------------------------------------------------------------
				-------------------- INSERE VALORES NA TABELA DE NOVA FIBRA -----------------------------
				-----------------------------------------------------------------------------------------
				


				DELETE FROM TBL_RE_BASE_NOVA_FIBRA WHERE ANOMES = @ANOMES AND TIPO_INDICADOR = 'REAL'

				INSERT INTO TBL_RE_BASE_NOVA_FIBRA

				SELECT /*DISTINCT*/ 'VL' AS INDBD ,DBO.FORMAT_DATE(DT_ENVIO_PEDIDO,'YYYYMM') AS ANOMES ,dbo.fn_Remover_Acentos(NO_MUNICIPIO) AS NO_MUNICIPIO ,NU_FILIAL ,LEFT(DT_ENVIO_PEDIDO,10) AS DATA ,NO_CANAL_PLANEJAMENTO ,NO_VELOCIDADE
				                ,NO_CLASSE_PRODUTO ,1 AS VALOR ,DS_SITUACAO_ORDEM ,NO_TIPO_MEIO_PAGAMENTO ,CD_PDV_SAP ,CD_OFERTA ,NO_OFERTA ,DT_ULTIMA_MODIFICACAO ,VL_TOTAL_RECORRENTE_OFERTA
					            ,MATRICULA_DO_VENDEDOR ,'REAL' AS TIPO_INDICADOR ,'VAREJO' AS SEGMENTO

				-- 07.06.22 RETIRADO O DISTINCT DA QUERY POIS FOI IDENTIFICADO QUE ESTAVAM EXPURGANDO VENDAS QUE NÃO DEVERIAM SER EXPURGADAS

				FROM TBL_CG_NOVA_FIBRA_VL

				WHERE IN_TESTE = '0' AND IN_BSIM = '0' AND DBO.FORMAT_DATE(DT_ENVIO_PEDIDO,'YYYYMMDD') <> DBO.FORMAT_DATE(GETDATE(),'YYYYMMDD') 

				UNION ALL

				SELECT /*DISTINCT*/ 'GROSS' AS INDBD ,DBO.FORMAT_DATE(DT_ATIVACAO,'YYYYMM') AS ANOMES ,dbo.fn_Remover_Acentos(NO_MUNICIPIO) AS NO_MUNICIPIO ,NU_FILIAL ,LEFT(DT_ATIVACAO,10) AS DATA ,NO_CANAL_PLANEJAMENTO ,NO_VELOCIDADE
								,NO_CLASSE_PRODUTO ,1 AS VALOR ,DS_SITUACAO_ORDEM ,NO_TIPO_MEIO_PAGAMENTO ,CD_PDV_SAP ,CD_OFERTA ,NO_OFERTA ,DT_ULTIMA_MODIFICACAO ,VL_TOTAL_RECORRENTE_OFERTA
								,MATRICULA_DO_VENDEDOR ,'REAL' AS TIPO_INDICADOR ,'VAREJO' AS SEGMENTO

				-- 07.06.22 RETIRADO O DISTINCT DA QUERY POIS FOI IDENTIFICADO QUE ESTAVAM EXPURGANDO VENDAS QUE NÃO DEVERIAM SER EXPURGADAS

				FROM TBL_CG_NOVA_FIBRA_GROSS

				WHERE IN_TESTE = '0' AND IN_BSIM = '0' AND DBO.FORMAT_DATE(DT_ATIVACAO,'YYYYMMDD') <> DBO.FORMAT_DATE(GETDATE(),'YYYYMMDD') 

				-- COMENTAR PARA NÃO MUDAR A TENDÊNCIA NAS CARGAS DAS PARCIAIS ....
				
				----------------------------------------------
				------EXECUTA CÁLCULO DE TENDÊNCIA (VLL)------
				----------------------------------------------

				EXEC sp_pc_tend_nova_fibra

				----------------------------------------------------------------------------------
				-----------------INÍCIO PROCESSO DE INSERÇÃO DA TENDÊNCIA-------------------------
				----------------------------------------------------------------------------------

				declare @MAIOR_DATA INTEGER
				declare @MAIOR_DATA_MES INTEGER

					select
							@MAIOR_DATA_MES = CONVERT(VARCHAR(8),DateAdd(Day,-1,Dateadd(Month,1, Convert(char(08),CONVERT(date, max(DATA)), 126)+'01')),112)
						  from TBL_RE_BaseResultadoDiario
						  WHERE LEFT(DATA,6) = @ANOMES AND GRUPO_PLANO = 'NOVA FIBRA' AND INDBD = 'VL'

					select
							@MAIOR_DATA = DBO.FORMAT_DATE(MAX(DATA),'YYYYMMDD')
						  from TBL_RE_BASE_NOVA_FIBRA
						  WHERE ANOMES = @ANOMES AND INDBD = 'VL' AND TIPO_INDICADOR = 'REAL'


					PRINT '-----------------------------------------------------------------------------------------'
					PRINT '-            INSERINDO TENDÊNCIA NAS TABELAS RE_RESULTADOS   MÊS: '+@ANOMES
					PRINT '-----------------------------------------------------------------------------------------'

					PRINT '-----------------------------------------------------------------------------------------'
					PRINT '-            VAREJO'
					PRINT '-----------------------------------------------------------------------------------------'

				--------------------------------
				--DROP/CREATE TABELA DE FATOR
				--------------------------------

				IF OBJECT_ID('TBL_PC_TEMP_FATOR_VAR_NOVA_TEND_FIBRA', 'U') IS NOT NULL
				DROP TABLE TBL_PC_TEMP_FATOR_VAR_NOVA_TEND_FIBRA


				SELECT * into TBL_PC_TEMP_FATOR_VAR_NOVA_TEND_FIBRA

				 FROM (
				
				SELECT DISTINCT B.NO_CANAL_PLANEJAMENTO,
								B.NU_FILIAL,
								B.NO_MUNICIPIO,
								CASE WHEN ((C.CANAL_RB IS NULL) OR (C.CANAL_RB IN ('S2S', 'Outros', 'TLV Outros'))) THEN 'Outros Nacionais'
								ELSE C.CANAL_RB END AS CANAL_RB
				FROM TBL_RE_BASE_NOVA_FIBRA AS B
				LEFT JOIN TBL_RE_DP_CanaisRelatorioBernardo AS C ON B.NO_CANAL_PLANEJAMENTO = C.CANAL_FINAL
				WHERE ANOMES = @ANOMES AND INDBD = 'VL' ) A

				LEFT JOIN (
				
				SELECT filial, municipio, canal, SUM(TEND)/sum(realizado) as fator FROM (
				
				SELECT data, dia_semana, filial, municipio, canal, qtd as tend,
					   case when realizado = 1 then qtd end as realizado
					   FROM tbl_pc_tend_nova_fibra ) TEND_NF group by filial, municipio, canal ) F

					ON A.CANAL_RB = F.canal AND
					   A.NU_FILIAL = F.filial AND
					   A.NO_MUNICIPIO = F.municipio

				------------------------------------------------------------------
				--IGUALANDO TENDÊNCIA AO REALIZADO NO CASO DE MÊS FECHADO
				------------------------------------------------------------------

					IF @MAIOR_DATA = @MAIOR_DATA_MES
						BEGIN
							UPDATE TBL_PC_TEMP_FATOR_VAR_NOVA_TEND_FIBRA SET FATOR = 1
							PRINT '----------------TENDÊNCIA IGUALADA AO REALIZADO----------------'
						END

				-------------------------------------
				--INSERT NA TBL_RE_BASE_NOVA_FIBRA
				-------------------------------------

				DELETE FROM [dbo].[TBL_RE_BASE_NOVA_FIBRA]
					WHERE TIPO_INDICADOR = 'TENDÊNCIA' AND ANOMES = @ANOMES
					AND INDBD = 'VL' AND SEGMENTO = 'VAREJO'

				INSERT INTO [dbo].[TBL_RE_BASE_NOVA_FIBRA]
				           ([INDBD]
							,[ANOMES]
							,[NO_MUNICIPIO]
							,[NU_FILIAL]
							,[DATA]
							,[NO_CANAL_PLANEJAMENTO]
							,[NO_VELOCIDADE]
							,[NO_CLASSE_PRODUTO]
							,[VALOR]
							,[DS_SITUACAO_ORDEM]
							,[NO_TIPO_MEIO_PAGAMENTO]
							,[CD_PDV_SAP]
							,[CD_OFERTA]
							,[NO_OFERTA]
							,[DT_ULTIMA_MODIFICACAO]
							,[vl_total_recorrente_oferta]
							,[matricula_do_vendedor]
							,[TIPO_INDICADOR]
							,[SEGMENTO]
							)
				    SELECT 
							[INDBD]
							,[ANOMES]
							,A.[NO_MUNICIPIO]
							,A.[NU_FILIAL]
							,[DATA]
							,A.[NO_CANAL_PLANEJAMENTO]
							,[NO_VELOCIDADE]
							,[NO_CLASSE_PRODUTO]
							,ISNULL((A.VALOR * F.fator),A.VALOR) AS [VALOR]
							,[DS_SITUACAO_ORDEM]
							,[NO_TIPO_MEIO_PAGAMENTO]
							,[CD_PDV_SAP]
							,[CD_OFERTA]
							,[NO_OFERTA]
							,[DT_ULTIMA_MODIFICACAO]
							,[vl_total_recorrente_oferta]
							,[matricula_do_vendedor]
							,'TENDÊNCIA' AS [TIPO_INDICADOR]
							,[SEGMENTO]
					FROM [dbo].[TBL_RE_BASE_NOVA_FIBRA] A
					LEFT JOIN TBL_PC_TEMP_FATOR_VAR_NOVA_TEND_FIBRA F ON
					   A.NO_CANAL_PLANEJAMENTO = F.NO_CANAL_PLANEJAMENTO AND
					   A.NU_FILIAL = F.filial AND
					   A.NO_MUNICIPIO = F.municipio
					WHERE TIPO_INDICADOR = 'REAL' AND ANOMES = @ANOMES
						AND INDBD = 'VL' AND SEGMENTO = 'VAREJO'

				---------------------------------------
				--DROP TABELA TEMPORÁRIA DE FATOR
				---------------------------------------

				IF OBJECT_ID('TBL_PC_TEMP_FATOR_VAR_NOVA_TEND_FIBRA', 'U') IS NOT NULL
				DROP TABLE TBL_PC_TEMP_FATOR_VAR_NOVA_TEND_FIBRA

				----------------------------------------------------------------------------------
				-------------------FIM PROCESSO DE INSERÇÃO DA TENDÊNCIA--------------------------
				----------------------------------------------------------------------------------
				
				-----------------------------------------------------------------------------------------
				--- INSERE VALORES NA BASE RESULTADOS - ADICIONADO POR NATÁLIA MOREIRA EM 30/11/2021 ----
				----------------------------------------------------------------------------------------- 

				DELETE  FROM TBL_RE_BASERESULTADOS
				WHERE  DATA = @ANOMES AND TIPO_INDICADOR IN ('REAL','TENDÊNCIA') AND GRUPO_PLANO = 'NOVA FIBRA'

				INSERT INTO TBL_RE_BASERESULTADOS

				SELECT INDBD , TIPO_INDICADOR ,ANOMES AS DATA ,NO_CANAL_PLANEJAMENTO AS CANAL_BOV ,NO_CANAL_PLANEJAMENTO AS CANAL_PARA ,CD_PDV_SAP AS COD_SAP ,NU_FILIAL AS FILIAL 
				       ,CASE WHEN B.DDD IS NULL THEN C.DDD ELSE B.DDD END AS DDD ,'NOVA FIBRA' AS GRUPO_PLANO ,'' AS PLANO ,'VA' AS SEGMENTO ,'FIBRA' AS PACOTE ,SUM(VALOR) AS VALOR 
					   ,NO_CANAL_PLANEJAMENTO AS CANAL_FINAL ,CD_OFERTA AS CAMPANHA ,'NA' HL ,MATRICULA_DO_VENDEDOR AS VENDEDOR ,CAST(VL_TOTAL_RECORRENTE_OFERTA AS FLOAT) AS ARPU 
					   ,'' AS PLANO_GERENCIAL ,'NI' AS ZONA_COMPETICAO ,'' AS PLANO_OFERTA ,'' AS PORTABILIDADE ,'' AS MULTIPRODUTO ,'ALONE' AS IND_COMBO ,'N' AS PEDIDO_UNICO

				FROM TBL_RE_BASE_NOVA_FIBRA A

				LEFT JOIN TBL_RE_DP_MUNICIPIO_DDD B
				ON A.NU_FILIAL = B.UF AND A.NO_MUNICIPIO = B.CIDADE

				LEFT JOIN TBL_RE_DP_FILIAL_DDD C
				ON A.NU_FILIAL = C.UF

				WHERE ANOMES = @ANOMES AND
					  SEGMENTO = 'VAREJO' AND
					  TIPO_INDICADOR IN ('REAL', 'TENDÊNCIA')

				GROUP BY INDBD ,TIPO_INDICADOR, ANOMES ,NO_CANAL_PLANEJAMENTO ,NO_CANAL_PLANEJAMENTO ,CD_PDV_SAP ,NU_FILIAL ,CASE WHEN B.DDD IS NULL THEN C.DDD ELSE B.DDD END ,NO_CANAL_PLANEJAMENTO ,CD_OFERTA 
						 ,MATRICULA_DO_VENDEDOR ,CAST(VL_TOTAL_RECORRENTE_OFERTA AS FLOAT)


				---------------------- INSERE O REAL COMO TENDÊNCIA (PARA O GROSS) --------------------------------------
				/*
				INSERT INTO TBL_RE_BASERESULTADOS 

				SELECT INDBD ,'TENDÊNCIA' AS TIPO_INDICADOR ,ANOMES AS DATA ,NO_CANAL_PLANEJAMENTO AS CANAL_BOV ,NO_CANAL_PLANEJAMENTO AS CANAL_PARA ,CD_PDV_SAP AS COD_SAP ,NU_FILIAL AS FILIAL 
				       ,CASE WHEN B.DDD IS NULL THEN C.DDD ELSE B.DDD END AS DDD ,'NOVA FIBRA' AS GRUPO_PLANO ,'' AS PLANO ,'VA' AS SEGMENTO ,'FIBRA' AS PACOTE ,SUM(VALOR) AS VALOR 
					   ,NO_CANAL_PLANEJAMENTO AS CANAL_FINAL ,CD_OFERTA AS CAMPANHA ,'NA' HL ,MATRICULA_DO_VENDEDOR AS VENDEDOR ,CAST(VL_TOTAL_RECORRENTE_OFERTA AS FLOAT) AS ARPU 
					   ,'' AS PLANO_GERENCIAL ,'NI' AS ZONA_COMPETICAO ,'' AS PLANO_OFERTA ,'' AS PORTABILIDADE ,'' AS MULTIPRODUTO ,'ALONE' AS IND_COMBO ,'N' AS PEDIDO_UNICO

				FROM TBL_RE_BASE_NOVA_FIBRA A

				LEFT JOIN TBL_RE_DP_MUNICIPIO_DDD B
				ON A.NU_FILIAL = B.UF AND A.NO_MUNICIPIO = B.CIDADE

				LEFT JOIN TBL_RE_DP_FILIAL_DDD C
				ON A.NU_FILIAL = C.UF

				WHERE ANOMES = @ANOMES AND
					  SEGMENTO = 'VAREJO' AND
					  TIPO_INDICADOR IN ('REAL') AND
					  INDBD = 'GROSS'

				GROUP BY INDBD ,ANOMES ,NO_CANAL_PLANEJAMENTO ,NO_CANAL_PLANEJAMENTO ,CD_PDV_SAP ,NU_FILIAL ,CASE WHEN B.DDD IS NULL THEN C.DDD ELSE B.DDD END ,NO_CANAL_PLANEJAMENTO ,CD_OFERTA 
					     ,MATRICULA_DO_VENDEDOR ,CAST(VL_TOTAL_RECORRENTE_OFERTA AS FLOAT)*/

				---------------------- INSERE VALORES NA BASE DIÁRIA -------------------------------------

				DELETE FROM DBO.TBL_RE_BASERESULTADODIARIO
				WHERE  LEFT(DATA,6) = @ANOMES AND GRUPO_PLANO = 'NOVA FIBRA'

				INSERT INTO TBL_RE_BASERESULTADODIARIO 

				SELECT INDBD ,REPLACE(DATA,'-','') AS DATA ,NO_CANAL_PLANEJAMENTO AS CANAL_BOV ,NO_CANAL_PLANEJAMENTO AS CANAL_PARA ,'' AS REGIONAL ,NU_FILIAL AS FILIAL 
					   ,CASE WHEN B.DDD IS NULL THEN C.DDD ELSE B.DDD END AS DDD ,CD_PDV_SAP AS COD_SAP ,'NOVA FIBRA' AS GRUPO_PLANO ,'' AS PLANO ,SUM(VALOR) AS QTD ,NO_CANAL_PLANEJAMENTO AS CANAL_FINAL 
					   ,'FIBRA' AS PACOTE ,CD_OFERTA AS CAMPANHA ,'NA' AS HL ,MATRICULA_DO_VENDEDOR AS VENDEDOR ,CAST(VL_TOTAL_RECORRENTE_OFERTA AS FLOAT) AS ARPU 
					   ,'' AS PLANO_GERENCIAL ,'NI' AS ZONA_COMPETICAO ,'' AS PORTABILIDADE ,'ALONE' AS IND_COMBO ,'N' AS PEDIDO_UNICO

				FROM TBL_RE_BASE_NOVA_FIBRA A

				LEFT JOIN TBL_RE_DP_MUNICIPIO_DDD B
				ON A.NU_FILIAL = B.UF AND A.NO_MUNICIPIO = B.CIDADE

				LEFT JOIN TBL_RE_DP_FILIAL_DDD C
				ON A.NU_FILIAL = C.UF

				WHERE ANOMES = @ANOMES AND TIPO_INDICADOR = 'REAL'

				GROUP BY INDBD ,REPLACE(DATA,'-','') ,NO_CANAL_PLANEJAMENTO ,NO_CANAL_PLANEJAMENTO ,NU_FILIAL ,CASE WHEN B.DDD IS NULL THEN C.DDD ELSE B.DDD END ,CD_PDV_SAP ,NO_CANAL_PLANEJAMENTO 
					     ,CD_OFERTA ,MATRICULA_DO_VENDEDOR ,CAST(VL_TOTAL_RECORRENTE_OFERTA AS FLOAT)


				EXEC [dbo].[SP_PC_CG_IND_Acompanhamento_Diario_Final] @ANOMES

				-----------------------------------------------------------------------------------------
				-------------------- REGISTRO DE FIM DE EXECUÇÃO -----------------------------
				-----------------------------------------------------------------------------------------
				INSERT INTO TBL_PC_TEMPO_PROCEDURES
				SELECT
				'SP_PC_NOVA_FIBRA' AS [PROCEDURE],
				'FIM' AS INI_FIM,
				GETDATE() AS DATA_HORA
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
    cursor.execute(comando_sql)
    conexao.commit()
    conexao.close()
    print('Conexão Fechada')

def coloca_comentario_procedure_nova_fibra_sql():
    comando_sql='''
				ALTER PROCEDURE [dbo].[SP_PC_NOVA_FIBRA] AS

				DECLARE @ANOMES AS VARCHAR(6)
				SET @ANOMES = (SELECT DISTINCT DBO.FORMAT_DATE(DT_MES,'YYYYMM') FROM TBL_CG_NOVA_FIBRA_VL)

				-----------------------------------------------------------------------------------------
				-------------------- REGISTRO DE INICIO DE EXECUÇÃO -----------------------------
				-----------------------------------------------------------------------------------------
				INSERT INTO TBL_PC_TEMPO_PROCEDURES
				SELECT
				'SP_PC_NOVA_FIBRA' AS [PROCEDURE],
				'INICIO' AS INI_FIM,
				GETDATE() AS DATA_HORA

				-----------------------------------------------------------------------------------------
				-------------------- INSERE VALORES NA TABELA DE NOVA FIBRA -----------------------------
				-----------------------------------------------------------------------------------------
				


				DELETE FROM TBL_RE_BASE_NOVA_FIBRA WHERE ANOMES = @ANOMES AND TIPO_INDICADOR = 'REAL'

				INSERT INTO TBL_RE_BASE_NOVA_FIBRA

				SELECT /*DISTINCT*/ 'VL' AS INDBD ,DBO.FORMAT_DATE(DT_ENVIO_PEDIDO,'YYYYMM') AS ANOMES ,dbo.fn_Remover_Acentos(NO_MUNICIPIO) AS NO_MUNICIPIO ,NU_FILIAL ,LEFT(DT_ENVIO_PEDIDO,10) AS DATA ,NO_CANAL_PLANEJAMENTO ,NO_VELOCIDADE
				                ,NO_CLASSE_PRODUTO ,1 AS VALOR ,DS_SITUACAO_ORDEM ,NO_TIPO_MEIO_PAGAMENTO ,CD_PDV_SAP ,CD_OFERTA ,NO_OFERTA ,DT_ULTIMA_MODIFICACAO ,VL_TOTAL_RECORRENTE_OFERTA
					            ,MATRICULA_DO_VENDEDOR ,'REAL' AS TIPO_INDICADOR ,'VAREJO' AS SEGMENTO

				-- 07.06.22 RETIRADO O DISTINCT DA QUERY POIS FOI IDENTIFICADO QUE ESTAVAM EXPURGANDO VENDAS QUE NÃO DEVERIAM SER EXPURGADAS

				FROM TBL_CG_NOVA_FIBRA_VL

				WHERE IN_TESTE = '0' AND IN_BSIM = '0' AND DBO.FORMAT_DATE(DT_ENVIO_PEDIDO,'YYYYMMDD') <> DBO.FORMAT_DATE(GETDATE(),'YYYYMMDD') 

				UNION ALL

				SELECT /*DISTINCT*/ 'GROSS' AS INDBD ,DBO.FORMAT_DATE(DT_ATIVACAO,'YYYYMM') AS ANOMES ,dbo.fn_Remover_Acentos(NO_MUNICIPIO) AS NO_MUNICIPIO ,NU_FILIAL ,LEFT(DT_ATIVACAO,10) AS DATA ,NO_CANAL_PLANEJAMENTO ,NO_VELOCIDADE
								,NO_CLASSE_PRODUTO ,1 AS VALOR ,DS_SITUACAO_ORDEM ,NO_TIPO_MEIO_PAGAMENTO ,CD_PDV_SAP ,CD_OFERTA ,NO_OFERTA ,DT_ULTIMA_MODIFICACAO ,VL_TOTAL_RECORRENTE_OFERTA
								,MATRICULA_DO_VENDEDOR ,'REAL' AS TIPO_INDICADOR ,'VAREJO' AS SEGMENTO

				-- 07.06.22 RETIRADO O DISTINCT DA QUERY POIS FOI IDENTIFICADO QUE ESTAVAM EXPURGANDO VENDAS QUE NÃO DEVERIAM SER EXPURGADAS

				FROM TBL_CG_NOVA_FIBRA_GROSS

				WHERE IN_TESTE = '0' AND IN_BSIM = '0' AND DBO.FORMAT_DATE(DT_ATIVACAO,'YYYYMMDD') <> DBO.FORMAT_DATE(GETDATE(),'YYYYMMDD') 

				-- COMENTAR PARA NÃO MUDAR A TENDÊNCIA NAS CARGAS DAS PARCIAIS ....
				/*
				----------------------------------------------
				------EXECUTA CÁLCULO DE TENDÊNCIA (VLL)------
				----------------------------------------------

				EXEC sp_pc_tend_nova_fibra

				----------------------------------------------------------------------------------
				-----------------INÍCIO PROCESSO DE INSERÇÃO DA TENDÊNCIA-------------------------
				----------------------------------------------------------------------------------

				declare @MAIOR_DATA INTEGER
				declare @MAIOR_DATA_MES INTEGER

					select
							@MAIOR_DATA_MES = CONVERT(VARCHAR(8),DateAdd(Day,-1,Dateadd(Month,1, Convert(char(08),CONVERT(date, max(DATA)), 126)+'01')),112)
						  from TBL_RE_BaseResultadoDiario
						  WHERE LEFT(DATA,6) = @ANOMES AND GRUPO_PLANO = 'NOVA FIBRA' AND INDBD = 'VL'

					select
							@MAIOR_DATA = DBO.FORMAT_DATE(MAX(DATA),'YYYYMMDD')
						  from TBL_RE_BASE_NOVA_FIBRA
						  WHERE ANOMES = @ANOMES AND INDBD = 'VL' AND TIPO_INDICADOR = 'REAL'


					PRINT '-----------------------------------------------------------------------------------------'
					PRINT '-            INSERINDO TENDÊNCIA NAS TABELAS RE_RESULTADOS   MÊS: '+@ANOMES
					PRINT '-----------------------------------------------------------------------------------------'

					PRINT '-----------------------------------------------------------------------------------------'
					PRINT '-            VAREJO'
					PRINT '-----------------------------------------------------------------------------------------'

				--------------------------------
				--DROP/CREATE TABELA DE FATOR
				--------------------------------

				IF OBJECT_ID('TBL_PC_TEMP_FATOR_VAR_NOVA_TEND_FIBRA', 'U') IS NOT NULL
				DROP TABLE TBL_PC_TEMP_FATOR_VAR_NOVA_TEND_FIBRA


				SELECT * into TBL_PC_TEMP_FATOR_VAR_NOVA_TEND_FIBRA

				 FROM (
				
				SELECT DISTINCT B.NO_CANAL_PLANEJAMENTO,
								B.NU_FILIAL,
								B.NO_MUNICIPIO,
								CASE WHEN ((C.CANAL_RB IS NULL) OR (C.CANAL_RB IN ('S2S', 'Outros', 'TLV Outros'))) THEN 'Outros Nacionais'
								ELSE C.CANAL_RB END AS CANAL_RB
				FROM TBL_RE_BASE_NOVA_FIBRA AS B
				LEFT JOIN TBL_RE_DP_CanaisRelatorioBernardo AS C ON B.NO_CANAL_PLANEJAMENTO = C.CANAL_FINAL
				WHERE ANOMES = @ANOMES AND INDBD = 'VL' ) A

				LEFT JOIN (
				
				SELECT filial, municipio, canal, SUM(TEND)/sum(realizado) as fator FROM (
				
				SELECT data, dia_semana, filial, municipio, canal, qtd as tend,
					   case when realizado = 1 then qtd end as realizado
					   FROM tbl_pc_tend_nova_fibra ) TEND_NF group by filial, municipio, canal ) F

					ON A.CANAL_RB = F.canal AND
					   A.NU_FILIAL = F.filial AND
					   A.NO_MUNICIPIO = F.municipio

				------------------------------------------------------------------
				--IGUALANDO TENDÊNCIA AO REALIZADO NO CASO DE MÊS FECHADO
				------------------------------------------------------------------

					IF @MAIOR_DATA = @MAIOR_DATA_MES
						BEGIN
							UPDATE TBL_PC_TEMP_FATOR_VAR_NOVA_TEND_FIBRA SET FATOR = 1
							PRINT '----------------TENDÊNCIA IGUALADA AO REALIZADO----------------'
						END

				-------------------------------------
				--INSERT NA TBL_RE_BASE_NOVA_FIBRA
				-------------------------------------

				DELETE FROM [dbo].[TBL_RE_BASE_NOVA_FIBRA]
					WHERE TIPO_INDICADOR = 'TENDÊNCIA' AND ANOMES = @ANOMES
					AND INDBD = 'VL' AND SEGMENTO = 'VAREJO'

				INSERT INTO [dbo].[TBL_RE_BASE_NOVA_FIBRA]
				           ([INDBD]
							,[ANOMES]
							,[NO_MUNICIPIO]
							,[NU_FILIAL]
							,[DATA]
							,[NO_CANAL_PLANEJAMENTO]
							,[NO_VELOCIDADE]
							,[NO_CLASSE_PRODUTO]
							,[VALOR]
							,[DS_SITUACAO_ORDEM]
							,[NO_TIPO_MEIO_PAGAMENTO]
							,[CD_PDV_SAP]
							,[CD_OFERTA]
							,[NO_OFERTA]
							,[DT_ULTIMA_MODIFICACAO]
							,[vl_total_recorrente_oferta]
							,[matricula_do_vendedor]
							,[TIPO_INDICADOR]
							,[SEGMENTO]
							)
				    SELECT 
							[INDBD]
							,[ANOMES]
							,A.[NO_MUNICIPIO]
							,A.[NU_FILIAL]
							,[DATA]
							,A.[NO_CANAL_PLANEJAMENTO]
							,[NO_VELOCIDADE]
							,[NO_CLASSE_PRODUTO]
							,ISNULL((A.VALOR * F.fator),A.VALOR) AS [VALOR]
							,[DS_SITUACAO_ORDEM]
							,[NO_TIPO_MEIO_PAGAMENTO]
							,[CD_PDV_SAP]
							,[CD_OFERTA]
							,[NO_OFERTA]
							,[DT_ULTIMA_MODIFICACAO]
							,[vl_total_recorrente_oferta]
							,[matricula_do_vendedor]
							,'TENDÊNCIA' AS [TIPO_INDICADOR]
							,[SEGMENTO]
					FROM [dbo].[TBL_RE_BASE_NOVA_FIBRA] A
					LEFT JOIN TBL_PC_TEMP_FATOR_VAR_NOVA_TEND_FIBRA F ON
					   A.NO_CANAL_PLANEJAMENTO = F.NO_CANAL_PLANEJAMENTO AND
					   A.NU_FILIAL = F.filial AND
					   A.NO_MUNICIPIO = F.municipio
					WHERE TIPO_INDICADOR = 'REAL' AND ANOMES = @ANOMES
						AND INDBD = 'VL' AND SEGMENTO = 'VAREJO'

				---------------------------------------
				--DROP TABELA TEMPORÁRIA DE FATOR
				---------------------------------------

				IF OBJECT_ID('TBL_PC_TEMP_FATOR_VAR_NOVA_TEND_FIBRA', 'U') IS NOT NULL
				DROP TABLE TBL_PC_TEMP_FATOR_VAR_NOVA_TEND_FIBRA

				----------------------------------------------------------------------------------
				-------------------FIM PROCESSO DE INSERÇÃO DA TENDÊNCIA--------------------------
				----------------------------------------------------------------------------------
				*/
				-----------------------------------------------------------------------------------------
				--- INSERE VALORES NA BASE RESULTADOS - ADICIONADO POR NATÁLIA MOREIRA EM 30/11/2021 ----
				----------------------------------------------------------------------------------------- 

				DELETE  FROM TBL_RE_BASERESULTADOS
				WHERE  DATA = @ANOMES AND TIPO_INDICADOR IN ('REAL','TENDÊNCIA') AND GRUPO_PLANO = 'NOVA FIBRA'

				INSERT INTO TBL_RE_BASERESULTADOS

				SELECT INDBD , TIPO_INDICADOR ,ANOMES AS DATA ,NO_CANAL_PLANEJAMENTO AS CANAL_BOV ,NO_CANAL_PLANEJAMENTO AS CANAL_PARA ,CD_PDV_SAP AS COD_SAP ,NU_FILIAL AS FILIAL 
				       ,CASE WHEN B.DDD IS NULL THEN C.DDD ELSE B.DDD END AS DDD ,'NOVA FIBRA' AS GRUPO_PLANO ,'' AS PLANO ,'VA' AS SEGMENTO ,'FIBRA' AS PACOTE ,SUM(VALOR) AS VALOR 
					   ,NO_CANAL_PLANEJAMENTO AS CANAL_FINAL ,CD_OFERTA AS CAMPANHA ,'NA' HL ,MATRICULA_DO_VENDEDOR AS VENDEDOR ,CAST(VL_TOTAL_RECORRENTE_OFERTA AS FLOAT) AS ARPU 
					   ,'' AS PLANO_GERENCIAL ,'NI' AS ZONA_COMPETICAO ,'' AS PLANO_OFERTA ,'' AS PORTABILIDADE ,'' AS MULTIPRODUTO ,'ALONE' AS IND_COMBO ,'N' AS PEDIDO_UNICO

				FROM TBL_RE_BASE_NOVA_FIBRA A

				LEFT JOIN TBL_RE_DP_MUNICIPIO_DDD B
				ON A.NU_FILIAL = B.UF AND A.NO_MUNICIPIO = B.CIDADE

				LEFT JOIN TBL_RE_DP_FILIAL_DDD C
				ON A.NU_FILIAL = C.UF

				WHERE ANOMES = @ANOMES AND
					  SEGMENTO = 'VAREJO' AND
					  TIPO_INDICADOR IN ('REAL', 'TENDÊNCIA')

				GROUP BY INDBD ,TIPO_INDICADOR, ANOMES ,NO_CANAL_PLANEJAMENTO ,NO_CANAL_PLANEJAMENTO ,CD_PDV_SAP ,NU_FILIAL ,CASE WHEN B.DDD IS NULL THEN C.DDD ELSE B.DDD END ,NO_CANAL_PLANEJAMENTO ,CD_OFERTA 
						 ,MATRICULA_DO_VENDEDOR ,CAST(VL_TOTAL_RECORRENTE_OFERTA AS FLOAT)


				---------------------- INSERE O REAL COMO TENDÊNCIA (PARA O GROSS) --------------------------------------
				/*
				INSERT INTO TBL_RE_BASERESULTADOS 

				SELECT INDBD ,'TENDÊNCIA' AS TIPO_INDICADOR ,ANOMES AS DATA ,NO_CANAL_PLANEJAMENTO AS CANAL_BOV ,NO_CANAL_PLANEJAMENTO AS CANAL_PARA ,CD_PDV_SAP AS COD_SAP ,NU_FILIAL AS FILIAL 
				       ,CASE WHEN B.DDD IS NULL THEN C.DDD ELSE B.DDD END AS DDD ,'NOVA FIBRA' AS GRUPO_PLANO ,'' AS PLANO ,'VA' AS SEGMENTO ,'FIBRA' AS PACOTE ,SUM(VALOR) AS VALOR 
					   ,NO_CANAL_PLANEJAMENTO AS CANAL_FINAL ,CD_OFERTA AS CAMPANHA ,'NA' HL ,MATRICULA_DO_VENDEDOR AS VENDEDOR ,CAST(VL_TOTAL_RECORRENTE_OFERTA AS FLOAT) AS ARPU 
					   ,'' AS PLANO_GERENCIAL ,'NI' AS ZONA_COMPETICAO ,'' AS PLANO_OFERTA ,'' AS PORTABILIDADE ,'' AS MULTIPRODUTO ,'ALONE' AS IND_COMBO ,'N' AS PEDIDO_UNICO

				FROM TBL_RE_BASE_NOVA_FIBRA A

				LEFT JOIN TBL_RE_DP_MUNICIPIO_DDD B
				ON A.NU_FILIAL = B.UF AND A.NO_MUNICIPIO = B.CIDADE

				LEFT JOIN TBL_RE_DP_FILIAL_DDD C
				ON A.NU_FILIAL = C.UF

				WHERE ANOMES = @ANOMES AND
					  SEGMENTO = 'VAREJO' AND
					  TIPO_INDICADOR IN ('REAL') AND
					  INDBD = 'GROSS'

				GROUP BY INDBD ,ANOMES ,NO_CANAL_PLANEJAMENTO ,NO_CANAL_PLANEJAMENTO ,CD_PDV_SAP ,NU_FILIAL ,CASE WHEN B.DDD IS NULL THEN C.DDD ELSE B.DDD END ,NO_CANAL_PLANEJAMENTO ,CD_OFERTA 
					     ,MATRICULA_DO_VENDEDOR ,CAST(VL_TOTAL_RECORRENTE_OFERTA AS FLOAT)*/

				---------------------- INSERE VALORES NA BASE DIÁRIA -------------------------------------

				DELETE FROM DBO.TBL_RE_BASERESULTADODIARIO
				WHERE  LEFT(DATA,6) = @ANOMES AND GRUPO_PLANO = 'NOVA FIBRA'

				INSERT INTO TBL_RE_BASERESULTADODIARIO 

				SELECT INDBD ,REPLACE(DATA,'-','') AS DATA ,NO_CANAL_PLANEJAMENTO AS CANAL_BOV ,NO_CANAL_PLANEJAMENTO AS CANAL_PARA ,'' AS REGIONAL ,NU_FILIAL AS FILIAL 
					   ,CASE WHEN B.DDD IS NULL THEN C.DDD ELSE B.DDD END AS DDD ,CD_PDV_SAP AS COD_SAP ,'NOVA FIBRA' AS GRUPO_PLANO ,'' AS PLANO ,SUM(VALOR) AS QTD ,NO_CANAL_PLANEJAMENTO AS CANAL_FINAL 
					   ,'FIBRA' AS PACOTE ,CD_OFERTA AS CAMPANHA ,'NA' AS HL ,MATRICULA_DO_VENDEDOR AS VENDEDOR ,CAST(VL_TOTAL_RECORRENTE_OFERTA AS FLOAT) AS ARPU 
					   ,'' AS PLANO_GERENCIAL ,'NI' AS ZONA_COMPETICAO ,'' AS PORTABILIDADE ,'ALONE' AS IND_COMBO ,'N' AS PEDIDO_UNICO

				FROM TBL_RE_BASE_NOVA_FIBRA A

				LEFT JOIN TBL_RE_DP_MUNICIPIO_DDD B
				ON A.NU_FILIAL = B.UF AND A.NO_MUNICIPIO = B.CIDADE

				LEFT JOIN TBL_RE_DP_FILIAL_DDD C
				ON A.NU_FILIAL = C.UF

				WHERE ANOMES = @ANOMES AND TIPO_INDICADOR = 'REAL'

				GROUP BY INDBD ,REPLACE(DATA,'-','') ,NO_CANAL_PLANEJAMENTO ,NO_CANAL_PLANEJAMENTO ,NU_FILIAL ,CASE WHEN B.DDD IS NULL THEN C.DDD ELSE B.DDD END ,CD_PDV_SAP ,NO_CANAL_PLANEJAMENTO 
					     ,CD_OFERTA ,MATRICULA_DO_VENDEDOR ,CAST(VL_TOTAL_RECORRENTE_OFERTA AS FLOAT)


				EXEC [dbo].[SP_PC_CG_IND_Acompanhamento_Diario_Final] @ANOMES

				-----------------------------------------------------------------------------------------
				-------------------- REGISTRO DE INICIO DE EXECUÇÃO -----------------------------
				-----------------------------------------------------------------------------------------
				INSERT INTO TBL_PC_TEMPO_PROCEDURES
				SELECT
				'SP_PC_NOVA_FIBRA' AS [PROCEDURE],
				'FIM' AS INI_FIM,
				GETDATE() AS DATA_HORA

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
    cursor.execute(comando_sql)
    conexao.commit()
    conexao.close()
    print('Conexão Fechada')

def rotina_se_ok ():
    print('\x1b[1;31;42m' + 'Datas dos arquivos 6163 e 6162 iguais a data de Hoje. Continuando o processo de tendência...'+ '\x1b[0m')
    # 2) Rodar procedure com codigo "descomentado"
    #print('Alterando a Procedure SP_PC_NOVA_FIBRA - descomentando o bloco de tendência')
    #tira_comentario_procedure_nova_fibra_sql()

    # 3) Rodar procedure SP_PC_NOVA_FIBRA - OK
    #print('Executando a Procedure SP_PC_NOVA_FIBRA')
    #executa_procedure_sql()    

    # 4) Rodar procedure com codigo comentado - OK
    #print('Alterando a Procedure SP_PC_NOVA_FIBRA - comentando o bloco de tendência')
    #coloca_comentario_procedure_nova_fibra_sql()

    # 5) Rodar query e colocar no excel - OK  |  # 6) salvar na rede - OK
    print('Montando Excel com tabela dinamica e salvando na rede')
    montaExcelTendVll()

    # 7) mandar por e-mail - OK
    print('Enviando e-mail')
    enviaEmaileAnexo()

    print('\x1b[1;32;41m' + 'CONCLUIDO' + '\x1b[0m')


# 1) Verificar se bases rodaram no MONITOR DE CARGA - OK
print('Verificando datas no Monitor de Cargas...')
fim, down, ini = puxa_dts_cargas()
dia6163 = (fim[0].split(' ')[0]) #pega somente a data do primeiro registro: 6163
dia6162 = (fim[1].split(' ')[0]) #pega somente a data do segundo registro: 6162

#dia6162 = '18/10/2022'
print(f'Data do 6163: {dia6163}. Dia do 6162:{dia6162}')

tentativas = 1
while (hoje != dia6163) or (hoje != dia6162):
    print('\x1b[1;32;41m' + 'Datas Diferentes. Esperando 5 minutos' + '\x1b[0m')
    time.sleep(60) #1minutos
    print('\x1b[1;32;41m' + 'Esperando 4 minutos' + '\x1b[0m')
    time.sleep(60) #1minutos
    print('\x1b[1;32;41m' + 'Esperando 3 minutos' + '\x1b[0m')
    time.sleep(60) #1minutos
    print('\x1b[1;32;41m' + 'Esperando 2 minutos' + '\x1b[0m')
    time.sleep(60) #1minutos
    print('\x1b[1;32;41m' + 'Esperando 1 minuto' + '\x1b[0m')
    time.sleep(60) #1minuto
    tentativas = tentativas + 1
    print(f'Verificando datas no Monitor de Cargas...Tentativa número: {tentativas}')
    fim, down, ini = puxa_dts_cargas()
    ia6163 = (fim[0].split(' ')[0]) #pega somente a data do primeiro registro: 6163
    dia6162 = (fim[1].split(' ')[0]) #pega somente a data do segundo registro: 6162
    print(f'Data do 6163: {dia6163}. Dia do 6162:{dia6162}')


if hoje == dia6163 == dia6162:
    rotina_se_ok()
	#tira_comentario_procedure_nova_fibra_sql()
    print('ROTINA OK')
else:
    print('\x1b[1;32;41m' + 'ESCAPE Datas Diferentes. Continuar esperando' + '\x1b[0m')
   
