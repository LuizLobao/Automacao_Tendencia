import shutil,os,time
import time
import win32com.client as win32
import pandas as pd
import segredos
import pyodbc
from telnetlib import theNULL
from datetime import date, datetime
from openpyxl import load_workbook
from playwright.sync_api import sync_playwright

hoje = datetime.today().strftime('%d/%m/%Y')
AAAAMMDD = datetime.today().strftime('%Y%m%d')


def data_mod_arquivo(caminho):   # Verificar data de Modificação antes de Realizar a Cópia do Arquivo
    modificado = time.strftime('%d/%m/%Y', time.gmtime(os.path.getmtime(caminho)))
    return (modificado)

def puxa_dts_cargas():   #Verifica data das bases no Monitor de Carga
	
    with sync_playwright() as p:
        navegador = p.chromium.launch(headless=True)
        pagina = navegador.new_page()
        pagina.goto("http://10.20.83.116/aplicacao/monitor/")
        
        datafim = {pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[10]/td[9]').text_content() : pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[10]/td[8]').text_content(),
                   pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[7]/td[9]').text_content() : pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[7]/td[8]').text_content(),
                   pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[9]/td[9]').text_content() : pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[9]/td[8]').text_content(),
                   pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[8]/td[9]').text_content() : pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[8]/td[8]').text_content(),
                   pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[12]/td[9]').text_content() : pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[12]/td[8]').text_content(),
                   pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[11]/td[9]').text_content() : pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[11]/td[8]').text_content(),   
                   pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[82]/td[9]').text_content() : pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[82]/td[8]').text_content(),
                   pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[83]/td[9]').text_content() : pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[83]/td[8]').text_content()
                   }
        navegador.close()
        return datafim

def copia_arquivo_renomeia():
    shutil.copy(r"Y:\\Demonstrativo Gross.xlsb", fr'S:\\Resultados\\01_Relatorio Diario\\1 - Base Eventos\\02 - TENDÊNCIA\\Insumos_Tendência\\Demonstrativo Gross_{AAAAMMDD}.xlsb')
    print('Arquivo copiado!')

def monta_tabdin_demonstrativo_gross():
    excel_file = f'S:\\Resultados\\01_Relatorio Diario\\1 - Base Eventos\\02 - TENDÊNCIA\\Insumos_Tendência\\Demonstrativo Gross_{AAAAMMDD}.xlsb'
    dest_filename = f'S:\\Resultados\\01_Relatorio Diario\\1 - Base Eventos\\02 - TENDÊNCIA\\Insumos_Tendência\\Demonstrativo Gross_{AAAAMMDD}.xlsx'

    df = pd.read_excel(excel_file, sheet_name='Database', engine='pyxlsb')
    pt_instalacao = df.query('TIPO == "INSTALACAO"').pivot_table(
                                                        values="PROJ", 
                                                        index=["UF"], 
                                                        columns="MERCADO", 
                                                        aggfunc=sum,
                                                        fill_value=0,
                                                        margins=True, margins_name="INSTALACAO",
                                                        )
    pt_migracao = df.query('TIPO == "MIGRACAO"').pivot_table(
                                                        values=["PROD","PROJ"], 
                                                        #index=["UF"], 
                                                        index="MERCADO", 
                                                        aggfunc=sum,
                                                        fill_value=0,
                                                        margins=True, margins_name="MIGRACAO",
                                                        )
    with pd.ExcelWriter(dest_filename) as writer:
        pt_instalacao.to_excel(writer, sheet_name="TabDin",startcol=0, startrow=0)
        pt_migracao.to_excel(writer, sheet_name="TabDin",startcol=6, startrow=0)

    print('Tabelas dinamicas concluidas!')

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
					PRINT '-			INSERINDO TENDÊNCIA NAS TABELAS RE_RESULTADOS   MÊS: '+@ANOMES
					PRINT '-----------------------------------------------------------------------------------------'

					PRINT '-----------------------------------------------------------------------------------------'
					PRINT '-			VAREJO'
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
	print("Conectado ao banco para alterar a procedure - retirar comentários")
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
					PRINT '-			INSERINDO TENDÊNCIA NAS TABELAS RE_RESULTADOS   MÊS: '+@ANOMES
					PRINT '-----------------------------------------------------------------------------------------'

					PRINT '-----------------------------------------------------------------------------------------'
					PRINT '-			VAREJO'
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
	print("Conectado ao banco para alterar a procedure - colocar comentários")
	cursor = conexao.cursor()
	cursor.execute(comando_sql)
	conexao.commit()
	conexao.close()
	print('Conexão Fechada')

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
	<p>Lobão, Luiz</p>
	"""
	anexo = (f'S:\\Resultados\\01_Relatorio Diario\\1 - Base Eventos\\02 - TENDÊNCIA\\Insumos_Tendência\\Tend_VLL_Nova_Fibra_{AAAAMMDD}.xlsx')
	email.Attachments.Add(anexo)

	email.Send()
	print("Email Enviado")

def ATIVAR_TEND_TABLEAU_teste_Jan22():
	comando_sql='''
                ALTER PROCEDURE [dbo].[SP_PC_CG_IND_Acompanhamento_Diario_Final] @ANOMES VARCHAR(6)
                WITH RECOMPILE
                AS
                --SET IMPLICIT_TRANSACTIONS ON
                BEGIN-- try 

                    INSERT INTO TBL_PC_TEMPO_PROCEDURES
                    SELECT
                    'SP_PC_CG_IND_Acompanhamento_Diario_Final' AS [PROCEDURE],
                    'INICIO' AS INI_FIM,
                    GETDATE() AS DATA_HORA
                    
                    
                    
                    
                    
                    PRINT '-----------------------------------------------------------------------------------------'
                    PRINT '-            INICIO CARGA SP_PC_CG_IND_Acompanhamento_Diario_Final   MÊS: '+@ANOMES
                    PRINT '-----------------------------------------------------------------------------------------'

                    DECLARE @ANOMES_M1 CHAR(6)
                    SET @ANOMES_M1 = dbo.format_date(DATEADD(MONTH, DATEDIFF(MONTH, 0 , GETDATE()-1)-1,0),'YYYYMM')


                    delete from TBL_IND_VAR_BASEMETA_PORDU where LEFT(data,6) = @ANOMES
                    /* ************** ATIVAR LINHA ABAIXO PARA TRAVAR TENDÊNCIA ************ */
                    --AND TIPO_INDICADOR <> 'TENDÊNCIA'

                    insert into TBL_IND_VAR_BASEMETA_PORDU
                    SELECT	
                                CASE WHEN (INDBD IN ('GROSS', 'MIGRACAO')) THEN 'GROSS' ELSE INDBD END INDBD,
                                DATA,
                                TIPO_INDICADOR,
                                CASE WHEN GRUPO_PLANO = 'OI GALERA PRÉ' THEN 'PRÉ-PAGO' ELSE GRUPO_PLANO END GRUPO_PLANO,
                                CASE WHEN ((COD_SAP = '1042793') AND (GRUPO_PLANO = 'CONTROLE_BOLETO')) THEN 'Smart Message'
                                ELSE C.CANAL_RB END CANAL_RB,
                                R.REGIONAL REGIONAL,
                                B.FILIAL,
                                'VAREJO' SEGMENTO,
                                SUM(VALOR) VALOR
                        FROM TBL_RE_BASERESULTADOS AS B
                        LEFT JOIN TBL_RE_DP_REGIONALRELATORIOBERNARDO AS R ON B.FILIAL = R.FILIAL
                        LEFT JOIN TBL_RE_DP_CANAISRELATORIOBERNARDO AS C ON B.CANAL_FINAL = C.CANAL_FINAL
                        /* ************** ATIVAR LINHA ABAIXO PARA TRAVAR TENDÊNCIA ************ */
                        --WHERE TIPO_INDICADOR IN ('META', 'FORECAST', 'ORÇAMENTO', 'RE-FCST') AND DATA = @ANOMES
                        /* ************** COMENTAR LINHA ACIMA E ATIVAR LINHA ABAIXO PARA DESTRAVAR TENDÊNCIA ************ */
                        WHERE TIPO_INDICADOR IN ('META', 'FORECAST', 'ORÇAMENTO', 'RE-FCST','TENDÊNCIA') AND DATA = @ANOMES
                        AND GRUPO_PLANO NOT IN ('CONECTADO', 'RESIDENCIAL', 'COMPLETO','1P','2P BL','3G','CONTROLE','CONTROLE_BOLETO-M1','OI TV PRE-PAGO','PACOTE_DADOS','TV ANTENEIROS','PÓS OCT','PÓS OCT PLANOS','CONTROLE_VOZ','BL+POS','PRÉ-M1','PÓS TOTAL','PÓS TOTAL PLANOS')
                        AND INDBD <> 'CANCELAMENTO' 
                        GROUP BY 
                                CASE WHEN (INDBD IN ('GROSS', 'MIGRACAO')) THEN 'GROSS' ELSE INDBD END,
                                TIPO_INDICADOR,
                                CASE WHEN GRUPO_PLANO = 'OI GALERA PRÉ' THEN 'PRÉ-PAGO' ELSE GRUPO_PLANO END,
                                DATA,
                                CASE WHEN ((COD_SAP = '1042793') AND (GRUPO_PLANO = 'CONTROLE_BOLETO')) THEN 'Smart Message'
                                ELSE C.CANAL_RB END,
                                R.REGIONAL,
                                B.FILIAL
                        UNION ALL
                        SELECT	
                                CASE WHEN (INDBD IN ('VL', 'MIGRACAO')) THEN 'VL' ELSE INDBD END INDBD,
                                DATA,
                                TIPO_INDICADOR,
                                GRUPO_PLANO,
                                C.CANAL_RB,
                                R.REGIONAL REGIONAL,
                                B.FILIAL,
                                'VAREJO' SEGMENTO,
                                SUM(VALOR) VALOR
                        FROM TBL_RE_BASERESULTADOS AS B
                        LEFT JOIN TBL_RE_DP_REGIONALRELATORIOBERNARDO AS R ON B.FILIAL = R.FILIAL
                        LEFT JOIN TBL_RE_DP_CANAISRELATORIOBERNARDO AS C ON B.CANAL_FINAL = C.CANAL_FINAL
                        /* ************** ATIVAR LINHA ABAIXO PARA TRAVAR TENDÊNCIA ************ */
                        --WHERE TIPO_INDICADOR IN ('META', 'FORECAST', 'ORÇAMENTO', 'RE-FCST') AND DATA = @ANOMES
                        /* ************** COMENTAR LINHA ACIMA E ATIVAR LINHA ABAIXO PARA DESTRAVAR TENDÊNCIA ************ */
                        WHERE TIPO_INDICADOR IN ('META', 'FORECAST', 'ORÇAMENTO', 'RE-FCST','TENDÊNCIA') AND  DATA = @ANOMES
                        AND GRUPO_PLANO IN ('CONECTADO', 'RESIDENCIAL', 'COMPLETO')
                        AND INDBD IN ('VL', 'MIGRACAO')
                        GROUP BY 
                                CASE WHEN (INDBD IN ('VL', 'MIGRACAO')) THEN 'VL' ELSE INDBD END,
                                TIPO_INDICADOR,
                                GRUPO_PLANO,
                                DATA,
                                C.CANAL_RB,
                                R.REGIONAL,
                                B.FILIAL
                        UNION ALL
                        SELECT	
                                CASE WHEN (INDBD IN ('GROSS', 'MIGRACAO')) THEN 'GROSS' ELSE INDBD END INDBD,
                                DATA,
                                TIPO_INDICADOR,
                                GRUPO_PLANO,
                                C.CANAL_RB,
                                R.REGIONAL REGIONAL,
                                B.FILIAL,
                                'VAREJO' SEGMENTO,
                                SUM(VALOR) VALOR
                        FROM TBL_RE_BASERESULTADOS AS B
                        LEFT JOIN TBL_RE_DP_REGIONALRELATORIOBERNARDO AS R ON B.FILIAL = R.FILIAL
                        LEFT JOIN TBL_RE_DP_CANAISRELATORIOBERNARDO AS C ON B.CANAL_FINAL = C.CANAL_FINAL
                        /* ************** ATIVAR LINHA ABAIXO PARA TRAVAR TENDÊNCIA ************ */
                        --WHERE TIPO_INDICADOR IN ('META', 'FORECAST', 'ORÇAMENTO', 'RE-FCST') AND DATA = @ANOMES
                        /* ************** COMENTAR LINHA ACIMA E ATIVAR LINHA ABAIXO PARA DESTRAVAR TENDÊNCIA ************ */
                        WHERE TIPO_INDICADOR IN ('META', 'FORECAST', 'ORÇAMENTO', 'RE-FCST','TENDÊNCIA') AND DATA = @ANOMES
                        AND GRUPO_PLANO IN ('CONECTADO', 'RESIDENCIAL', 'COMPLETO')
                        AND INDBD IN ('GROSS', 'MIGRACAO')
                        GROUP BY 
                                CASE WHEN (INDBD IN ('GROSS', 'MIGRACAO')) THEN 'GROSS' ELSE INDBD END,
                                TIPO_INDICADOR,
                                GRUPO_PLANO,
                                DATA,
                                C.CANAL_RB,
                                R.REGIONAL,
                                B.FILIAL
                    UNION ALL
                    SELECT	
                                CASE WHEN (INDBD IN ('GROSS', 'MIGRACAO')) THEN 'GROSS' ELSE INDBD END INDBD,
                                DATA,
                                TIPO_INDICADOR,
                                GRUPO_PLANO,
                                C.CANAL_RB,
                                R.REGIONAL REGIONAL,
                                B.UF_CLIENTE FILIAL,
                                'EMP CLI' SEGMENTO,
                                SUM(VALOR) VALOR
                        FROM TBL_RE_BaseResultados_Empresarial AS B
                        LEFT JOIN TBL_RE_DP_REGIONALRELATORIOBERNARDO AS R ON B.UF_CLIENTE = R.FILIAL
                        LEFT JOIN TBL_RE_DP_CanaisRelatorioBernardo_Empresarial AS C ON B.CANAL_FINAL = C.CANAL_FINAL
                        /* ************** ATIVAR LINHA ABAIXO PARA TRAVAR TENDÊNCIA ************ */
                        --WHERE TIPO_INDICADOR IN ('META', 'FORECAST', 'ORÇAMENTO', 'RE-FCST') AND DATA = @ANOMES
                        /* ************** COMENTAR LINHA ACIMA E ATIVAR LINHA ABAIXO PARA DESTRAVAR TENDÊNCIA ************ */
                        WHERE TIPO_INDICADOR IN ('META', 'FORECAST', 'ORÇAMENTO', 'RE-FCST','TENDÊNCIA') AND DATA = @ANOMES
                        AND INDBD <> 'CANCELAMENTO' 
                        GROUP BY 
                                CASE WHEN (INDBD IN ('GROSS', 'MIGRACAO')) THEN 'GROSS' ELSE INDBD END,
                                TIPO_INDICADOR,
                                GRUPO_PLANO,
                                DATA,
                                C.CANAL_RB,
                                R.REGIONAL,
                                B.UF_CLIENTE
                        UNION ALL
                        SELECT	
                                CASE WHEN (INDBD IN ('GROSS', 'MIGRACAO')) THEN 'GROSS' ELSE INDBD END INDBD,
                                DATA,
                                TIPO_INDICADOR,
                                GRUPO_PLANO,
                                C.CANAL_RB,
                                R.REGIONAL REGIONAL,
                                B.UF_CARTEIRA FILIAL,
                                'EMP PDV' SEGMENTO,
                                SUM(VALOR) VALOR
                        FROM TBL_RE_BaseResultados_Empresarial AS B
                        LEFT JOIN TBL_RE_DP_REGIONALRELATORIOBERNARDO AS R ON B.UF_CARTEIRA = R.FILIAL
                        LEFT JOIN TBL_RE_DP_CanaisRelatorioBernardo_Empresarial AS C ON B.CANAL_FINAL = C.CANAL_FINAL
                        /* ************** ATIVAR LINHA ABAIXO PARA TRAVAR TENDÊNCIA ************ */
                        --WHERE TIPO_INDICADOR IN ('META', 'FORECAST', 'ORÇAMENTO', 'RE-FCST') AND DATA = @ANOMES
                        /* ************** COMENTAR LINHA ACIMA E ATIVAR LINHA ABAIXO PARA DESTRAVAR TENDÊNCIA ************ */ 
                        WHERE TIPO_INDICADOR IN ('META', 'FORECAST', 'ORÇAMENTO', 'RE-FCST','TENDÊNCIA') AND DATA = @ANOMES
                        AND INDBD <> 'CANCELAMENTO' 
                        AND REGIONAL_AGRUPADA IS NOT NULL
                        GROUP BY 
                                CASE WHEN (INDBD IN ('GROSS', 'MIGRACAO')) THEN 'GROSS' ELSE INDBD END,
                                TIPO_INDICADOR,
                                GRUPO_PLANO,
                                DATA,
                                C.CANAL_RB,
                                R.REGIONAL,
                                B.UF_CARTEIRA



                    DELETE FROM TBL_IND_VAR_ACOMPANHAMENTO_RESIDENCIAL_BASE WHERE DATA = @ANOMES
                    AND NOT (TIPO_INDICADOR = 'META' AND DATA IN ('202007', '202008', '202009') AND GRUPO_PLANO IN ('CONTROLE_BOLETO',
                    'CONTROLE_CARTAO', 'PÓS ALONE', 'PÓS OIT', 'PRÉ-PAGO', 'PRÉ-D3', 'MOVEL'))

                    insert into TBL_IND_VAR_ACOMPANHAMENTO_RESIDENCIAL_BASE
                    select 
                            TIPO_INDICADOR,
                            data,
                            INDBD,
                            GRUPO_PLANO,
                            CANAL, 
                            REGIONAL,
                            SEGMENTO,
                            CASE 
                                WHEN CANAL IN ('Smart Message','TLV Outros','Outros', 'Outros BRI','Condominios','','NULL','S2S','Outros EMP','S2S EMP','TLV Outros EMP' ) THEN 'Outros Nacionais' 
                                WHEN CANAL IN ('TLV Receptivo','TLV Ativo','TLV Ativo EMP','TLV Receptivo EMP' ) THEN 'TLV' 
                                WHEN CANAL IN ('WEB','WEB EMP') THEN 'WEB' 
                                WHEN CANAL = 'Anteneiros' THEN 'Anteneiros' 
                                WHEN CANAL IN ('TLV Receptivo BRI','TLV Outros BRI','TLV Ativo BRI','TLV BRI' ) THEN 'TLV BRI' 
                                WHEN CANAL LIKE '%EMP%' AND [SEGMENTO] <> 'VAREJO' THEN 'Gestão Regional EMP'
                                WHEN CANAL NOT LIKE '%EMP%' AND [SEGMENTO] <> 'VAREJO' THEN 'Gestão Regional VAR'
                            ELSE 'Gestão Regional' END GESTAO,
                            UF,
                            cast( DIA as varchar(2)) DIA,
                            META_DU QTD


                    from(	SELECT		distinct 
                            TIPO_INDICADOR,
                            a.data,
                            A.indbd INDBD,
                            CASE WHEN A.grupo_plano IN ('OI GALERA PRÉ', 'PRÉ-PAGO') THEN 'PRÉ-PAGO' ELSE A.grupo_plano END GRUPO_PLANO,
                            A.canal_rb CANAL, 
                            REG.REGIONAL_AGRUPADA REGIONAL,
                            SEGMENTO,
                            LTRIM(RTRIM(A.FILIAL)) UF,
                            '' AS DIA,
                            a.VALOR META_DU
                            -- ,A.VALOR
                            --,DU_MES_PRODUTO.VALOR DU_MES_PRODUTO
                            -- ,DU_MES.VALOR DU_MES
                            -- ,P.VALOR P
                        
                    FROM TBL_IND_VAR_BASEMETA_PORDU A
                    LEFT JOIN TBL_RE_DP_RegionalRelatorioBernardo AS REG ON A.FILIAL = REG.FILIAL
                                    
                    WHERE --a.canal_rb not in ('outros bri', 'tlv ativo bri', 'tlv receptivo bri', 'tlv outros bri','Anteneiros','tlv bri','bri') and
                                a.grupo_plano not in (	'CONECTADO', 'RESIDENCIAL', 'COMPLETO','1P','2P BL','3G','CONTROLE','CONTROLE_BOLETO-M1','OI TV PRE-PAGO','PACOTE_DADOS',
                                                            'TV ANTENEIROS','PÓS OCT','PÓS OCT PLANOS','CONTROLE_VOZ','BL+POS','PRÉ-M1','PÓS TOTAL','PÓS TOTAL PLANOS',
                                                            'BANDA LARGA FIBRA C','BANDA LARGA FIBRA EO','BANDA LARGA FIBRA EX OBRA','BANDA LARGA FIBRA SI','FIBRA C',
                                                            'FIBRA EO','FIBRA EX OBRA','FIBRA SI','FIXO FIBRA C','FIXO FIBRA EO','FIXO FIBRA EX OBRA','FIXO FIBRA SI',
                                                            'OI TV FIBRA C','OI TV FIBRA EO','OI TV FIBRA EX OBRA','OI TV FIBRA SI','OIT COMERCIAL',
                                                            'OIT CONECTADO','OIT SOLUCAO COMPLETA','VADA')
                            AND A.TIPO_INDICADOR = 'TENDÊNCIA'
                            AND A.DATA = @ANOMES

                            UNION ALL


                    SELECT		distinct 
                            TIPO_INDICADOR,
                            a.data,
                            A.indbd INDBD,
                            CASE WHEN A.grupo_plano IN ('OI GALERA PRÉ', 'PRÉ-PAGO') THEN 'PRÉ-PAGO' ELSE A.grupo_plano END GRUPO_PLANO,
                            A.canal_rb CANAL, 
                            REG.REGIONAL_AGRUPADA REGIONAL,
                            SEGMENTO,
                            LTRIM(RTRIM(A.FILIAL)) UF,
                            RIGHT(C.ANOMESDIA,2) AS DIA,
                            case 
                                when A.VALOR = 0 then 0
                                when DU_MES.VALOR > 0 or DU_MES.VALOR is not null then (A.VALOR/DU_MES.VALOR)*C.VALOR
                                when DU_MES_PRODUTO.VALOR > 0 or DU_MES_PRODUTO.VALOR is not null then (A.VALOR/DU_MES_PRODUTO.VALOR)*p.VALOR 
                            else	0 end META_DU
                            -- ,A.VALOR
                            --,DU_MES_PRODUTO.VALOR DU_MES_PRODUTO
                            -- ,DU_MES.VALOR DU_MES
                            -- ,P.VALOR P
                        
                    FROM TBL_IND_VAR_BASEMETA_PORDU A
                    LEFT JOIN TBL_RE_DP_RegionalRelatorioBernardo AS REG ON A.FILIAL = REG.FILIAL

                    /* TOTAL DU POR REGIONAL - CANAL */
                    left JOIN (
                                    SELECT 
                                        * 
                                    FROM TBL_IND_VAR_DU_SUMARIZADO
                                    where ANOMES = @ANOMES
                            ) DU_MES ON A.INDBD = DU_MES.INDBD AND 
                                                        --a.GRUPO_PLANO = DU_MES.PRODUTO_NOVO AND
                                                        case when A.grupo_plano = 'NOVA FIBRA' then 'FIBRA' ELSE A.GRUPO_PLANO END = DU_MES.PRODUTO_NOVO AND
                                                        CASE WHEN A.CANAL_RB = 'Smart Message' THEN 'TLV Outros' 
                                                        ELSE A.CANAL_RB END = DU_MES.CANAL AND
                                                        A.REGIONAL = DU_MES.REGIONAL AND
                                                        A.DATA = DU_MES.ANOMES

                    /* DU POR DIA PARA REGIONAL - CANAL */
                    left JOIN (SELECT ANOMESDIA,
                                    left(ANOMESDIA,6) ANOMES,
                                    INDBD,
                                    CASE WHEN DU_DIA.PRODUTO IN ('OI GALERA PRÉ', 'PRÉ-PAGO') THEN 'PRÉ-PAGO' ELSE DU_DIA.PRODUTO END PRODUTO_NOVO,
                                    CANAL,
                                    REGIONAL,
                                    valor
                                FROM TBL_pc_du AS DU_DIA
                                            where left(ANOMESDIA,6) = @ANOMES
                                ) C ON 
                                        A.INDBD = C.INDBD AND 
                                        case when A.grupo_plano = 'NOVA FIBRA' then 'FIBRA' ELSE A.GRUPO_PLANO END = C.PRODUTO_NOVO AND
                                        CASE WHEN A.CANAL_RB = 'Smart Message' THEN 'TLV Outros' ELSE
                                        A.CANAL_RB END = C.CANAL AND
                                        A.REGIONAL = C.REGIONAL and
                                        A.DATA = C.ANOMES
                    
                    /* TOTAL DU POR PRODUTOS */
                    left JOIN (
                                    SELECT 
                                            LEFT(ANOMESDIA,6) ANOMES,
                                            INDBD,
                                            PRODUTO_NOVO,
                                            SUM(cast(rtrim(ltrim(replace(VALOR,',','.'))) as float)) VALOR 
                                    FROM TBL_PC_DU_PRODUTO
                                    where left(ANOMESDIA,6) = @ANOMES
                                    group by 
                                            LEFT(ANOMESDIA,6),
                                            INDBD,
                                            PRODUTO_NOVO
                            ) DU_MES_PRODUTO ON	A.INDBD = DU_MES_PRODUTO.INDBD AND 
                                                    case when A.grupo_plano = 'NOVA FIBRA' then 'FIBRA' ELSE A.GRUPO_PLANO END = DU_MES_PRODUTO.PRODUTO_NOVO AND
                                                    A.DATA = DU_MES_PRODUTO.ANOMES

                    /* DU POR DIA PARA PRODUTO */
                    left JOIN (SELECT 
                                        ANOMESDIA,
                                        LEFT(ANOMESDIA,6) ANOMES,
                                        INDBD,
                                        PRODUTO_NOVO,
                                        SUM(cast(rtrim(ltrim(replace(VALOR,',','.'))) as float)) VALOR 
                                FROM TBL_PC_DU_PRODUTO
                                where left(ANOMESDIA,6) = @ANOMES
                                group by 
                                        ANOMESDIA,
                                        LEFT(ANOMESDIA,6),
                                        INDBD,
                                        PRODUTO_NOVO
                                ) P ON 
                                        A.INDBD = P.INDBD AND 
                                        case when A.grupo_plano = 'NOVA FIBRA' then 'FIBRA' ELSE A.GRUPO_PLANO END = P.PRODUTO_NOVO AND
                                        A.DATA = P.ANOMES
                    
                                    
                    WHERE --a.canal_rb not in ('outros bri', 'tlv ativo bri', 'tlv receptivo bri', 'tlv outros bri','Anteneiros','tlv bri','bri') and
                                a.grupo_plano not in (	'CONECTADO', 'RESIDENCIAL', 'COMPLETO','1P','2P BL','3G','CONTROLE','CONTROLE_BOLETO-M1','OI TV PRE-PAGO','PACOTE_DADOS',
                                                            'TV ANTENEIROS','PÓS OCT','PÓS OCT PLANOS','CONTROLE_VOZ','BL+POS','PRÉ-M1','PÓS TOTAL','PÓS TOTAL PLANOS',
                                                            'BANDA LARGA FIBRA C','BANDA LARGA FIBRA EO','BANDA LARGA FIBRA EX OBRA','BANDA LARGA FIBRA SI','FIBRA C',
                                                            'FIBRA EO','FIBRA EX OBRA','FIBRA SI','FIXO FIBRA C','FIXO FIBRA EO','FIXO FIBRA EX OBRA','FIXO FIBRA SI',
                                                            'OI TV FIBRA C','OI TV FIBRA EO','OI TV FIBRA EX OBRA','OI TV FIBRA SI','OIT COMERCIAL',
                                                            'OIT CONECTADO','OIT SOLUCAO COMPLETA','VADA')
                            AND A.TIPO_INDICADOR <> 'TENDÊNCIA'
                            AND A.DATA = @ANOMES
                            AND NOT (TIPO_INDICADOR = 'META' AND DATA IN ('202007', '202008', '202009') AND GRUPO_PLANO IN ('CONTROLE_BOLETO',
                            'CONTROLE_CARTAO', 'PÓS ALONE', 'PÓS OIT', 'PRÉ-PAGO', 'PRÉ-D3', 'MOVEL'))

                    union all

                    SELECT  'REAL' TIPO_INDICADOR,
                            left(data,6) anomes,
                            INDBD,	
                        GRUPO_PLANO,
                        CASE WHEN ((B.COD_SAP = '1042793') AND (GRUPO_PLANO = 'CONTROLE_BOLETO')) THEN 'Smart Message'
                        ELSE C.CANAL_RB END canal,
                        R.REGIONAL_AGRUPADA regional,
                        'VAREJO' SEGMENTO,
                        LTRIM(RTRIM(B.FILIAL)) AS UF,
                        RIGHT(DATA,2) AS DIA,
                        SUM(B.QTD) AS valor

                    FROM TBL_RE_BaseResultadoDiario B	
                    LEFT JOIN TBL_RE_DP_RegionalRelatorioBernardo AS R ON B.FILIAL = R.FILIAL	
                    LEFT JOIN TBL_RE_DP_CanaisRelatorioBernardo AS C ON B.CANAL_FINAL = C.CANAL_FINAL	

                    WHERE --c.canal_rb not in ('outros bri', 'tlv ativo bri', 'tlv receptivo bri', 'tlv outros bri','Anteneiros','tlv bri','bri') and
                    b.grupo_plano not in ('CONECTADO', 'RESIDENCIAL', 'COMPLETO','1P','2P BL','3G','CONTROLE','CONTROLE_BOLETO-M1','OI TV PRE-PAGO','PACOTE_DADOS','TV ANTENEIROS','PÓS OCT','PÓS OCT PLANOS','CONTROLE_VOZ','BL+POS','PRÉ-M1','PÓS TOTAL','PÓS TOTAL PLANOS')
                    AND B.INDBD <> 'CANCELAMENTO' 
                    AND LEFT(DATA,6) = @ANOMES
                    
                    GROUP BY INDBD,	
                        GRUPO_PLANO,
                        left(data,6),
                        CASE WHEN ((B.COD_SAP = '1042793') AND (GRUPO_PLANO = 'CONTROLE_BOLETO')) THEN 'Smart Message'
                        ELSE C.CANAL_RB END,
                        R.REGIONAL_AGRUPADA,
                    
                        LTRIM(RTRIM(B.FILIAL)),
                        RIGHT(DATA,2)

                    union all

                    SELECT  'REAL' TIPO_INDICADOR,
                            left(data,6) anomes,
                            INDBD,	
                        GRUPO_PLANO,
                        C.CANAL_RB canal,
                        R.REGIONAL_AGRUPADA regional,
                        'EMP CLI' SEGMENTO,
                        LTRIM(RTRIM(B.UF_CLIENTE)) AS UF,
                        RIGHT(DATA,2) AS DIA,
                        SUM(B.QTD) AS valor

                    FROM TBL_RE_BaseResultadoDiario_Empresarial B	
                    LEFT JOIN TBL_RE_DP_RegionalRelatorioBernardo AS R ON B.UF_CLIENTE = R.FILIAL	
                    LEFT JOIN TBL_RE_DP_CanaisRelatorioBernardo_Empresarial AS C ON B.CANAL_FINAL = C.CANAL_FINAL	

                    WHERE --c.canal_rb not in ('outros bri', 'tlv ativo bri', 'tlv receptivo bri', 'tlv outros bri','Anteneiros','tlv bri','bri') and
                    b.grupo_plano not in ('CONECTADO', 'RESIDENCIAL', 'COMPLETO','1P','2P BL','3G','CONTROLE','CONTROLE_BOLETO-M1','OI TV PRE-PAGO','PACOTE_DADOS','TV ANTENEIROS','PÓS OCT','PÓS OCT PLANOS','CONTROLE_VOZ','BL+POS','PRÉ-M1','PÓS TOTAL','PÓS TOTAL PLANOS')
                    AND B.INDBD <> 'CANCELAMENTO' 
                    AND LEFT(DATA,6)  = @ANOMES
                    
                    GROUP BY INDBD,	
                        GRUPO_PLANO,
                        left(data,6),
                        C.CANAL_RB,
                        R.REGIONAL_AGRUPADA,
                    
                        LTRIM(RTRIM(B.UF_CLIENTE)),
                        RIGHT(DATA,2)

                    union all

                    SELECT  'REAL' TIPO_INDICADOR,
                            left(data,6) anomes,
                            INDBD,	
                        GRUPO_PLANO,
                        C.CANAL_RB canal,
                        R.REGIONAL_AGRUPADA regional,
                        'EMP PDV' SEGMENTO,
                        LTRIM(RTRIM(B.UF_CARTEIRA)) AS UF,
                        RIGHT(DATA,2) AS DIA,
                        SUM(B.QTD) AS valor
                    FROM TBL_RE_BaseResultadoDiario_Empresarial B	
                    LEFT JOIN TBL_RE_DP_RegionalRelatorioBernardo AS R ON B.UF_CARTEIRA = R.FILIAL	
                    LEFT JOIN TBL_RE_DP_CanaisRelatorioBernardo_Empresarial AS C ON B.CANAL_FINAL = C.CANAL_FINAL	

                    WHERE --c.canal_rb not in ('outros bri', 'tlv ativo bri', 'tlv receptivo bri', 'tlv outros bri','Anteneiros','tlv bri','bri') and
                    b.grupo_plano not in ('CONECTADO', 'RESIDENCIAL', 'COMPLETO','1P','2P BL','3G','CONTROLE','CONTROLE_BOLETO-M1','OI TV PRE-PAGO','PACOTE_DADOS','TV ANTENEIROS','PÓS OCT','PÓS OCT PLANOS','CONTROLE_VOZ','BL+POS','PRÉ-M1','PÓS TOTAL','PÓS TOTAL PLANOS')
                    AND B.INDBD <> 'CANCELAMENTO' 
                    AND LEFT(DATA,6) = @ANOMES
                    
                    GROUP BY INDBD,	
                        GRUPO_PLANO,
                        left(data,6),
                        C.CANAL_RB,
                        R.REGIONAL_AGRUPADA,
                        LTRIM(RTRIM(B.UF_CARTEIRA)),
                        RIGHT(DATA,2)
                    ) t
                    /*
                    TRUNCATE TABLE TBL_IND_VAR_ACOMPANHAMENTO_RESIDENCIAL_BASE_NOVA

                    DECLARE @DATA_2 AS VARCHAR (6)

                    SET @DATA_2 = (SELECT MAX(DATA) FROM TBL_IND_VAR_ACOMPANHAMENTO_RESIDENCIAL_BASE)

                    insert into TBL_IND_VAR_ACOMPANHAMENTO_RESIDENCIAL_BASE_NOVA
                SELECT 
                    --TIPO_INDICADOR,
                    DATA,
                    INDBD,
                    case 
                                    when GRUPO_PLANO in ('CONTROLE BOLETO','CONTROLE_BOLETO') then 'CONTROLE_BOLETO' 
                                    when GRUPO_PLANO in ('CONTROLE CARTAO','CONTROLE_CARTAO') then 'CONTROLE_CARTAO' 
                                    when GRUPO_PLANO in ('MOVEL','MÓVEL') then 'MOVEL' 
                                    when GRUPO_PLANO in ('OI GALERA PRÉ','PRÉ-PAGO') then 'PRÉ-PAGO' 
                            else GRUPO_PLANO end GRUPO_PLANO,
                    CANAL,
                    REGIONAL,
                    SEGMENTO,
                    GESTAO GESTÃO,
                    UF,
                    DIA,
                    0 AS COMPROMISSO,
                    SUM(CASE WHEN  TIPO_INDICADOR = 'FORECAST' THEN QTD ELSE 0 END) AS FORECAST,
                    SUM(CASE WHEN  TIPO_INDICADOR = 'META' THEN QTD ELSE 0 END) AS META,
                    SUM(CASE WHEN  TIPO_INDICADOR = 'ORÇAMENTO' THEN QTD ELSE 0 END) AS ORCAMENTO,
                    SUM(CASE WHEN  TIPO_INDICADOR = 'REAL' THEN QTD ELSE 0 END) AS REAL,
                    SUM(CASE WHEN  TIPO_INDICADOR = 'TENDÊNCIA' THEN QTD ELSE 0 END) AS TENDENCIA

                FROM TBL_IND_VAR_ACOMPANHAMENTO_RESIDENCIAL_BASE
                /* REMOVENDO AS VISÕES DE FIBRA EM OBRA - SOLICITAÇÃO DO MARIO - 07/05/2019 */
                WHERE GRUPO_PLANO NOT LIKE '%C' AND GRUPO_PLANO NOT LIKE '%EO' AND GRUPO_PLANO NOT LIKE '%SI' AND GRUPO_PLANO NOT LIKE '%EX OBRA'

                GROUP BY 	DATA,
                    INDBD,
                    GRUPO_PLANO,
                    CANAL,
                    REGIONAL,
                    SEGMENTO,
                    GESTAO,
                    UF,
                    DIA

                UNION ALL

                select 
                    --TIPO_INDICADOR,
                    ANOMES DATA,
                    INDBD,
                    case 
                                    when GRUPO_PLANO in ('CONTROLE BOLETO','CONTROLE_BOLETO') then 'CONTROLE_BOLETO' 
                                    when GRUPO_PLANO in ('CONTROLE CARTAO','CONTROLE_CARTAO') then 'CONTROLE_CARTAO' 
                                    when GRUPO_PLANO in ('MOVEL','MÓVEL') then 'MOVEL' 
                                    when GRUPO_PLANO in ('OI GALERA PRÉ','PRÉ-PAGO') then 'PRÉ-PAGO' 
                            else GRUPO_PLANO end GRUPO_PLANO,
                    CANAL,
                    REGIONAL,
                    CASE WHEN UNIDADE_NEGOCIO = 'EMPRESARIAL' THEN 'EMP PDV' ELSE UNIDADE_NEGOCIO END AS SEGMENTO,
                    GESTÃO,
                    UF,
                    DIA,
                    SUM(QTD) AS COMPROMISSO,
                    0 FORECAST,
                    0 META,
                    0 ORCAMENTO,
                    0 REAL,
                    0 TENDENCIA

                from TBL_PC_BASEMETA_RELATORIO_MANOEL

                WHERE TIPO_INDICADOR = 'COMPROMISSO'
                AND ANOMES >= @DATA_2

                GROUP BY 	--TIPO_INDICADOR,
                    ANOMES,
                    INDBD,
                    GRUPO_PLANO,
                    CANAL,
                    REGIONAL,
                    UNIDADE_NEGOCIO,
                    GESTÃO,
                    UF,
                    DIA*/


                    INSERT INTO TBL_PC_TEMPO_PROCEDURES
                    SELECT
                    'SP_PC_CG_IND_Acompanhamento_Diario_Final' AS [PROCEDURE],
                    'FIM' AS INI_FIM,
                    GETDATE() AS DATA_HORA
                    
                    
                    
                    
                    print '--- Commit ---'	
                    PRINT '-----------------------------------------------------------------------------------------'
                    PRINT '-            FIM CARGA SP_PC_CG_IND_Acompanhamento_Diario_Final   MÊS: '+@ANOMES
                    PRINT '-----------------------------------------------------------------------------------------'
                    
                    
                    --COMMIT
                end
				'''

	dados_conexao = (
		"Driver={SQL Server};"
		f"Server={segredos.db_server};"
		f"Database={segredos.db_name};"
		f"UID={segredos.db_user};"
		f"PWD={segredos.db_pass}"
	)
	conexao = pyodbc.connect(dados_conexao)
	print("Conectado ao banco para alterar a procedure - retirar comentários")
	cursor = conexao.cursor()
	cursor.execute(comando_sql)
	conexao.commit()
	conexao.close()
	print('Conexão Fechada')

def ATIVAR_TEND_TABLEAU_teste_Jan22_somenteFibra():
	comando_sql='''
				ALTER PROCEDURE [dbo].[SP_PC_CG_IND_Acompanhamento_Diario_Final] @ANOMES VARCHAR(6)
                WITH RECOMPILE
                AS
                --SET IMPLICIT_TRANSACTIONS ON
                BEGIN-- try 

                    INSERT INTO TBL_PC_TEMPO_PROCEDURES
                    SELECT
                    'SP_PC_CG_IND_Acompanhamento_Diario_Final' AS [PROCEDURE],
                    'INICIO' AS INI_FIM,
                    GETDATE() AS DATA_HORA





                    PRINT '-----------------------------------------------------------------------------------------'
                    PRINT '-            INICIO CARGA SP_PC_CG_IND_Acompanhamento_Diario_Final   MÊS: '+@ANOMES
                    PRINT '-----------------------------------------------------------------------------------------'

                    DECLARE @ANOMES_M1 CHAR(6)
                    SET @ANOMES_M1 = dbo.format_date(DATEADD(MONTH, DATEDIFF(MONTH, 0 , GETDATE()-1)-1,0),'YYYYMM')


                    delete from TBL_IND_VAR_BASEMETA_PORDU where LEFT(data,6) = @ANOMES
                    /* ************** ATIVAR LINHA ABAIXO PARA TRAVAR TENDÊNCIA ************ */
                    --AND TIPO_INDICADOR <> 'TENDÊNCIA'
                    AND NOT (TIPO_INDICADOR = 'TENDÊNCIA' AND GRUPO_PLANO NOT LIKE '%FIBRA')

                    insert into TBL_IND_VAR_BASEMETA_PORDU
                    SELECT	
                                CASE WHEN (INDBD IN ('GROSS', 'MIGRACAO')) THEN 'GROSS' ELSE INDBD END INDBD,
                                DATA,
                                TIPO_INDICADOR,
                                CASE WHEN GRUPO_PLANO = 'OI GALERA PRÉ' THEN 'PRÉ-PAGO' ELSE GRUPO_PLANO END GRUPO_PLANO,
                                CASE WHEN ((COD_SAP = '1042793') AND (GRUPO_PLANO = 'CONTROLE_BOLETO')) THEN 'Smart Message'
                                ELSE C.CANAL_RB END CANAL_RB,
                                R.REGIONAL REGIONAL,
                                B.FILIAL,
                                'VAREJO' SEGMENTO,
                                SUM(VALOR) VALOR
                        FROM TBL_RE_BASERESULTADOS AS B
                        LEFT JOIN TBL_RE_DP_REGIONALRELATORIOBERNARDO AS R ON B.FILIAL = R.FILIAL
                        LEFT JOIN TBL_RE_DP_CANAISRELATORIOBERNARDO AS C ON B.CANAL_FINAL = C.CANAL_FINAL
                        /* ************** ATIVAR LINHA ABAIXO PARA TRAVAR TENDÊNCIA ************ */
                        --WHERE TIPO_INDICADOR IN ('META', 'FORECAST', 'ORÇAMENTO', 'RE-FCST') AND DATA = @ANOMES
                        /* ************** COMENTAR LINHA ACIMA E ATIVAR LINHA ABAIXO PARA DESTRAVAR TENDÊNCIA ************ */
                        WHERE TIPO_INDICADOR IN ('META', 'FORECAST', 'ORÇAMENTO', 'RE-FCST','TENDÊNCIA') AND DATA = @ANOMES
                        AND GRUPO_PLANO NOT IN ('CONECTADO', 'RESIDENCIAL', 'COMPLETO','1P','2P BL','3G','CONTROLE','CONTROLE_BOLETO-M1','OI TV PRE-PAGO','PACOTE_DADOS','TV ANTENEIROS','PÓS OCT','PÓS OCT PLANOS','CONTROLE_VOZ','BL+POS','PRÉ-M1','PÓS TOTAL','PÓS TOTAL PLANOS')
                        AND INDBD <> 'CANCELAMENTO'
                        AND NOT (TIPO_INDICADOR = 'TENDÊNCIA' AND GRUPO_PLANO NOT LIKE '%FIBRA')
                        GROUP BY 
                                CASE WHEN (INDBD IN ('GROSS', 'MIGRACAO')) THEN 'GROSS' ELSE INDBD END,
                                TIPO_INDICADOR,
                                CASE WHEN GRUPO_PLANO = 'OI GALERA PRÉ' THEN 'PRÉ-PAGO' ELSE GRUPO_PLANO END,
                                DATA,
                                CASE WHEN ((COD_SAP = '1042793') AND (GRUPO_PLANO = 'CONTROLE_BOLETO')) THEN 'Smart Message'
                                ELSE C.CANAL_RB END,
                                R.REGIONAL,
                                B.FILIAL
                        UNION ALL
                        SELECT	
                                CASE WHEN (INDBD IN ('VL', 'MIGRACAO')) THEN 'VL' ELSE INDBD END INDBD,
                                DATA,
                                TIPO_INDICADOR,
                                GRUPO_PLANO,
                                C.CANAL_RB,
                                R.REGIONAL REGIONAL,
                                B.FILIAL,
                                'VAREJO' SEGMENTO,
                                SUM(VALOR) VALOR
                        FROM TBL_RE_BASERESULTADOS AS B
                        LEFT JOIN TBL_RE_DP_REGIONALRELATORIOBERNARDO AS R ON B.FILIAL = R.FILIAL
                        LEFT JOIN TBL_RE_DP_CANAISRELATORIOBERNARDO AS C ON B.CANAL_FINAL = C.CANAL_FINAL
                        /* ************** ATIVAR LINHA ABAIXO PARA TRAVAR TENDÊNCIA ************ */
                        --WHERE TIPO_INDICADOR IN ('META', 'FORECAST', 'ORÇAMENTO', 'RE-FCST') AND DATA = @ANOMES
                        /* ************** COMENTAR LINHA ACIMA E ATIVAR LINHA ABAIXO PARA DESTRAVAR TENDÊNCIA ************ */
                        WHERE TIPO_INDICADOR IN ('META', 'FORECAST', 'ORÇAMENTO', 'RE-FCST','TENDÊNCIA') AND  DATA = @ANOMES
                        AND GRUPO_PLANO IN ('CONECTADO', 'RESIDENCIAL', 'COMPLETO')
                        AND INDBD IN ('VL', 'MIGRACAO')
                        GROUP BY 
                                CASE WHEN (INDBD IN ('VL', 'MIGRACAO')) THEN 'VL' ELSE INDBD END,
                                TIPO_INDICADOR,
                                GRUPO_PLANO,
                                DATA,
                                C.CANAL_RB,
                                R.REGIONAL,
                                B.FILIAL
                        UNION ALL
                        SELECT	
                                CASE WHEN (INDBD IN ('GROSS', 'MIGRACAO')) THEN 'GROSS' ELSE INDBD END INDBD,
                                DATA,
                                TIPO_INDICADOR,
                                GRUPO_PLANO,
                                C.CANAL_RB,
                                R.REGIONAL REGIONAL,
                                B.FILIAL,
                                'VAREJO' SEGMENTO,
                                SUM(VALOR) VALOR
                        FROM TBL_RE_BASERESULTADOS AS B
                        LEFT JOIN TBL_RE_DP_REGIONALRELATORIOBERNARDO AS R ON B.FILIAL = R.FILIAL
                        LEFT JOIN TBL_RE_DP_CANAISRELATORIOBERNARDO AS C ON B.CANAL_FINAL = C.CANAL_FINAL
                        /* ************** ATIVAR LINHA ABAIXO PARA TRAVAR TENDÊNCIA ************ */
                        --WHERE TIPO_INDICADOR IN ('META', 'FORECAST', 'ORÇAMENTO', 'RE-FCST') AND DATA = @ANOMES
                        /* ************** COMENTAR LINHA ACIMA E ATIVAR LINHA ABAIXO PARA DESTRAVAR TENDÊNCIA ************ */
                        WHERE TIPO_INDICADOR IN ('META', 'FORECAST', 'ORÇAMENTO', 'RE-FCST','TENDÊNCIA') AND DATA = @ANOMES
                        AND GRUPO_PLANO IN ('CONECTADO', 'RESIDENCIAL', 'COMPLETO')
                        AND INDBD IN ('GROSS', 'MIGRACAO')
                        GROUP BY 
                                CASE WHEN (INDBD IN ('GROSS', 'MIGRACAO')) THEN 'GROSS' ELSE INDBD END,
                                TIPO_INDICADOR,
                                GRUPO_PLANO,
                                DATA,
                                C.CANAL_RB,
                                R.REGIONAL,
                                B.FILIAL
                    UNION ALL
                    SELECT	
                                CASE WHEN (INDBD IN ('GROSS', 'MIGRACAO')) THEN 'GROSS' ELSE INDBD END INDBD,
                                DATA,
                                TIPO_INDICADOR,
                                GRUPO_PLANO,
                                C.CANAL_RB,
                                R.REGIONAL REGIONAL,
                                B.UF_CLIENTE FILIAL,
                                'EMP CLI' SEGMENTO,
                                SUM(VALOR) VALOR
                        FROM TBL_RE_BaseResultados_Empresarial AS B
                        LEFT JOIN TBL_RE_DP_REGIONALRELATORIOBERNARDO AS R ON B.UF_CLIENTE = R.FILIAL
                        LEFT JOIN TBL_RE_DP_CanaisRelatorioBernardo_Empresarial AS C ON B.CANAL_FINAL = C.CANAL_FINAL
                        /* ************** ATIVAR LINHA ABAIXO PARA TRAVAR TENDÊNCIA ************ */
                        --WHERE TIPO_INDICADOR IN ('META', 'FORECAST', 'ORÇAMENTO', 'RE-FCST') AND DATA = @ANOMES
                        /* ************** COMENTAR LINHA ACIMA E ATIVAR LINHA ABAIXO PARA DESTRAVAR TENDÊNCIA ************ */
                        WHERE TIPO_INDICADOR IN ('META', 'FORECAST', 'ORÇAMENTO', 'RE-FCST','TENDÊNCIA') AND DATA = @ANOMES
                        AND INDBD <> 'CANCELAMENTO' 
                        AND NOT (TIPO_INDICADOR = 'TENDÊNCIA' AND GRUPO_PLANO NOT LIKE '%FIBRA')
                        GROUP BY 
                                CASE WHEN (INDBD IN ('GROSS', 'MIGRACAO')) THEN 'GROSS' ELSE INDBD END,
                                TIPO_INDICADOR,
                                GRUPO_PLANO,
                                DATA,
                                C.CANAL_RB,
                                R.REGIONAL,
                                B.UF_CLIENTE
                        UNION ALL
                        SELECT	
                                CASE WHEN (INDBD IN ('GROSS', 'MIGRACAO')) THEN 'GROSS' ELSE INDBD END INDBD,
                                DATA,
                                TIPO_INDICADOR,
                                GRUPO_PLANO,
                                C.CANAL_RB,
                                R.REGIONAL REGIONAL,
                                B.UF_CARTEIRA FILIAL,
                                'EMP PDV' SEGMENTO,
                                SUM(VALOR) VALOR
                        FROM TBL_RE_BaseResultados_Empresarial AS B
                        LEFT JOIN TBL_RE_DP_REGIONALRELATORIOBERNARDO AS R ON B.UF_CARTEIRA = R.FILIAL
                        LEFT JOIN TBL_RE_DP_CanaisRelatorioBernardo_Empresarial AS C ON B.CANAL_FINAL = C.CANAL_FINAL
                        /* ************** ATIVAR LINHA ABAIXO PARA TRAVAR TENDÊNCIA ************ */
                        --WHERE TIPO_INDICADOR IN ('META', 'FORECAST', 'ORÇAMENTO', 'RE-FCST') AND DATA = @ANOMES
                        /* ************** COMENTAR LINHA ACIMA E ATIVAR LINHA ABAIXO PARA DESTRAVAR TENDÊNCIA ************ */ 
                        WHERE TIPO_INDICADOR IN ('META', 'FORECAST', 'ORÇAMENTO', 'RE-FCST','TENDÊNCIA') AND DATA = @ANOMES
                        AND INDBD <> 'CANCELAMENTO' 
                        AND REGIONAL_AGRUPADA IS NOT NULL
                        AND NOT (TIPO_INDICADOR = 'TENDÊNCIA' AND GRUPO_PLANO NOT LIKE '%FIBRA')
                        GROUP BY 
                                CASE WHEN (INDBD IN ('GROSS', 'MIGRACAO')) THEN 'GROSS' ELSE INDBD END,
                                TIPO_INDICADOR,
                                GRUPO_PLANO,
                                DATA,
                                C.CANAL_RB,
                                R.REGIONAL,
                                B.UF_CARTEIRA



                    DELETE FROM TBL_IND_VAR_ACOMPANHAMENTO_RESIDENCIAL_BASE WHERE DATA = @ANOMES
                    AND NOT (TIPO_INDICADOR = 'META' AND DATA IN ('202007', '202008', '202009') AND GRUPO_PLANO IN ('CONTROLE_BOLETO',
                    'CONTROLE_CARTAO', 'PÓS ALONE', 'PÓS OIT', 'PRÉ-PAGO', 'PRÉ-D3', 'MOVEL'))

                    insert into TBL_IND_VAR_ACOMPANHAMENTO_RESIDENCIAL_BASE
                    select 
                            TIPO_INDICADOR,
                            data,
                            INDBD,
                            GRUPO_PLANO,
                            CANAL, 
                            REGIONAL,
                            SEGMENTO,
                            CASE 
                                WHEN CANAL IN ('Smart Message','TLV Outros','Outros', 'Outros BRI','Condominios','','NULL','S2S','Outros EMP','S2S EMP','TLV Outros EMP' ) THEN 'Outros Nacionais' 
                                WHEN CANAL IN ('TLV Receptivo','TLV Ativo','TLV Ativo EMP','TLV Receptivo EMP' ) THEN 'TLV' 
                                WHEN CANAL IN ('WEB','WEB EMP') THEN 'WEB' 
                                WHEN CANAL = 'Anteneiros' THEN 'Anteneiros' 
                                WHEN CANAL IN ('TLV Receptivo BRI','TLV Outros BRI','TLV Ativo BRI','TLV BRI' ) THEN 'TLV BRI' 
                                WHEN CANAL LIKE '%EMP%' AND [SEGMENTO] <> 'VAREJO' THEN 'Gestão Regional EMP'
                                WHEN CANAL NOT LIKE '%EMP%' AND [SEGMENTO] <> 'VAREJO' THEN 'Gestão Regional VAR'
                            ELSE 'Gestão Regional' END GESTAO,
                            UF,
                            cast( DIA as varchar(2)) DIA,
                            META_DU QTD


                    from(	SELECT		distinct 
                            TIPO_INDICADOR,
                            a.data,
                            A.indbd INDBD,
                            CASE WHEN A.grupo_plano IN ('OI GALERA PRÉ', 'PRÉ-PAGO') THEN 'PRÉ-PAGO' ELSE A.grupo_plano END GRUPO_PLANO,
                            A.canal_rb CANAL, 
                            REG.REGIONAL_AGRUPADA REGIONAL,
                            SEGMENTO,
                            LTRIM(RTRIM(A.FILIAL)) UF,
                            '' AS DIA,
                            a.VALOR META_DU
                            -- ,A.VALOR
                            --,DU_MES_PRODUTO.VALOR DU_MES_PRODUTO
                            -- ,DU_MES.VALOR DU_MES
                            -- ,P.VALOR P
                        
                    FROM TBL_IND_VAR_BASEMETA_PORDU A
                    LEFT JOIN TBL_RE_DP_RegionalRelatorioBernardo AS REG ON A.FILIAL = REG.FILIAL
                                    
                    WHERE --a.canal_rb not in ('outros bri', 'tlv ativo bri', 'tlv receptivo bri', 'tlv outros bri','Anteneiros','tlv bri','bri') and
                                a.grupo_plano not in (	'CONECTADO', 'RESIDENCIAL', 'COMPLETO','1P','2P BL','3G','CONTROLE','CONTROLE_BOLETO-M1','OI TV PRE-PAGO','PACOTE_DADOS',
                                                            'TV ANTENEIROS','PÓS OCT','PÓS OCT PLANOS','CONTROLE_VOZ','BL+POS','PRÉ-M1','PÓS TOTAL','PÓS TOTAL PLANOS',
                                                            'BANDA LARGA FIBRA C','BANDA LARGA FIBRA EO','BANDA LARGA FIBRA EX OBRA','BANDA LARGA FIBRA SI','FIBRA C',
                                                            'FIBRA EO','FIBRA EX OBRA','FIBRA SI','FIXO FIBRA C','FIXO FIBRA EO','FIXO FIBRA EX OBRA','FIXO FIBRA SI',
                                                            'OI TV FIBRA C','OI TV FIBRA EO','OI TV FIBRA EX OBRA','OI TV FIBRA SI','OIT COMERCIAL',
                                                            'OIT CONECTADO','OIT SOLUCAO COMPLETA','VADA')
                            AND A.TIPO_INDICADOR = 'TENDÊNCIA'
                            AND A.DATA = @ANOMES

                            UNION ALL


                    SELECT		distinct 
                            TIPO_INDICADOR,
                            a.data,
                            A.indbd INDBD,
                            CASE WHEN A.grupo_plano IN ('OI GALERA PRÉ', 'PRÉ-PAGO') THEN 'PRÉ-PAGO' ELSE A.grupo_plano END GRUPO_PLANO,
                            A.canal_rb CANAL, 
                            REG.REGIONAL_AGRUPADA REGIONAL,
                            SEGMENTO,
                            LTRIM(RTRIM(A.FILIAL)) UF,
                            RIGHT(C.ANOMESDIA,2) AS DIA,
                            case 
                                when A.VALOR = 0 then 0
                                when DU_MES.VALOR > 0 or DU_MES.VALOR is not null then (A.VALOR/DU_MES.VALOR)*C.VALOR
                                when DU_MES_PRODUTO.VALOR > 0 or DU_MES_PRODUTO.VALOR is not null then (A.VALOR/DU_MES_PRODUTO.VALOR)*p.VALOR 
                            else	0 end META_DU
                            -- ,A.VALOR
                            --,DU_MES_PRODUTO.VALOR DU_MES_PRODUTO
                            -- ,DU_MES.VALOR DU_MES
                            -- ,P.VALOR P
                        
                    FROM TBL_IND_VAR_BASEMETA_PORDU A
                    LEFT JOIN TBL_RE_DP_RegionalRelatorioBernardo AS REG ON A.FILIAL = REG.FILIAL

                    /* TOTAL DU POR REGIONAL - CANAL */
                    left JOIN (
                                    SELECT 
                                        * 
                                    FROM TBL_IND_VAR_DU_SUMARIZADO
                                    where ANOMES = @ANOMES
                            ) DU_MES ON A.INDBD = DU_MES.INDBD AND 
                                                        --a.GRUPO_PLANO = DU_MES.PRODUTO_NOVO AND
                                                        case when A.grupo_plano = 'NOVA FIBRA' then 'FIBRA' ELSE A.GRUPO_PLANO END = DU_MES.PRODUTO_NOVO AND
                                                        CASE WHEN A.CANAL_RB = 'Smart Message' THEN 'TLV Outros' 
                                                        ELSE A.CANAL_RB END = DU_MES.CANAL AND
                                                        A.REGIONAL = DU_MES.REGIONAL AND
                                                        A.DATA = DU_MES.ANOMES

                    /* DU POR DIA PARA REGIONAL - CANAL */
                    left JOIN (SELECT ANOMESDIA,
                                    left(ANOMESDIA,6) ANOMES,
                                    INDBD,
                                    CASE WHEN DU_DIA.PRODUTO IN ('OI GALERA PRÉ', 'PRÉ-PAGO') THEN 'PRÉ-PAGO' ELSE DU_DIA.PRODUTO END PRODUTO_NOVO,
                                    CANAL,
                                    REGIONAL,
                                    valor
                                FROM TBL_pc_du AS DU_DIA
                                            where left(ANOMESDIA,6) = @ANOMES
                                ) C ON 
                                        A.INDBD = C.INDBD AND 
                                        case when A.grupo_plano = 'NOVA FIBRA' then 'FIBRA' ELSE A.GRUPO_PLANO END = C.PRODUTO_NOVO AND
                                        CASE WHEN A.CANAL_RB = 'Smart Message' THEN 'TLV Outros' ELSE
                                        A.CANAL_RB END = C.CANAL AND
                                        A.REGIONAL = C.REGIONAL and
                                        A.DATA = C.ANOMES
                    
                    /* TOTAL DU POR PRODUTOS */
                    left JOIN (
                                    SELECT 
                                            LEFT(ANOMESDIA,6) ANOMES,
                                            INDBD,
                                            PRODUTO_NOVO,
                                            SUM(cast(rtrim(ltrim(replace(VALOR,',','.'))) as float)) VALOR 
                                    FROM TBL_PC_DU_PRODUTO
                                    where left(ANOMESDIA,6) = @ANOMES
                                    group by 
                                            LEFT(ANOMESDIA,6),
                                            INDBD,
                                            PRODUTO_NOVO
                            ) DU_MES_PRODUTO ON	A.INDBD = DU_MES_PRODUTO.INDBD AND 
                                                    case when A.grupo_plano = 'NOVA FIBRA' then 'FIBRA' ELSE A.GRUPO_PLANO END = DU_MES_PRODUTO.PRODUTO_NOVO AND
                                                    A.DATA = DU_MES_PRODUTO.ANOMES

                    /* DU POR DIA PARA PRODUTO */
                    left JOIN (SELECT 
                                        ANOMESDIA,
                                        LEFT(ANOMESDIA,6) ANOMES,
                                        INDBD,
                                        PRODUTO_NOVO,
                                        SUM(cast(rtrim(ltrim(replace(VALOR,',','.'))) as float)) VALOR 
                                FROM TBL_PC_DU_PRODUTO
                                where left(ANOMESDIA,6) = @ANOMES
                                group by 
                                        ANOMESDIA,
                                        LEFT(ANOMESDIA,6),
                                        INDBD,
                                        PRODUTO_NOVO
                                ) P ON 
                                        A.INDBD = P.INDBD AND 
                                        case when A.grupo_plano = 'NOVA FIBRA' then 'FIBRA' ELSE A.GRUPO_PLANO END = P.PRODUTO_NOVO AND
                                        A.DATA = P.ANOMES
                    
                                    
                    WHERE --a.canal_rb not in ('outros bri', 'tlv ativo bri', 'tlv receptivo bri', 'tlv outros bri','Anteneiros','tlv bri','bri') and
                                a.grupo_plano not in (	'CONECTADO', 'RESIDENCIAL', 'COMPLETO','1P','2P BL','3G','CONTROLE','CONTROLE_BOLETO-M1','OI TV PRE-PAGO','PACOTE_DADOS',
                                                            'TV ANTENEIROS','PÓS OCT','PÓS OCT PLANOS','CONTROLE_VOZ','BL+POS','PRÉ-M1','PÓS TOTAL','PÓS TOTAL PLANOS',
                                                            'BANDA LARGA FIBRA C','BANDA LARGA FIBRA EO','BANDA LARGA FIBRA EX OBRA','BANDA LARGA FIBRA SI','FIBRA C',
                                                            'FIBRA EO','FIBRA EX OBRA','FIBRA SI','FIXO FIBRA C','FIXO FIBRA EO','FIXO FIBRA EX OBRA','FIXO FIBRA SI',
                                                            'OI TV FIBRA C','OI TV FIBRA EO','OI TV FIBRA EX OBRA','OI TV FIBRA SI','OIT COMERCIAL',
                                                            'OIT CONECTADO','OIT SOLUCAO COMPLETA','VADA')
                            AND A.TIPO_INDICADOR <> 'TENDÊNCIA'
                            AND A.DATA = @ANOMES
                            AND NOT (TIPO_INDICADOR = 'META' AND DATA IN ('202007', '202008', '202009') AND GRUPO_PLANO IN ('CONTROLE_BOLETO',
                            'CONTROLE_CARTAO', 'PÓS ALONE', 'PÓS OIT', 'PRÉ-PAGO', 'PRÉ-D3', 'MOVEL'))

                    union all

                    SELECT  'REAL' TIPO_INDICADOR,
                            left(data,6) anomes,
                            INDBD,	
                        GRUPO_PLANO,
                        CASE WHEN ((B.COD_SAP = '1042793') AND (GRUPO_PLANO = 'CONTROLE_BOLETO')) THEN 'Smart Message'
                        ELSE C.CANAL_RB END canal,
                        R.REGIONAL_AGRUPADA regional,
                        'VAREJO' SEGMENTO,
                        LTRIM(RTRIM(B.FILIAL)) AS UF,
                        RIGHT(DATA,2) AS DIA,
                        SUM(B.QTD) AS valor

                    FROM TBL_RE_BaseResultadoDiario B	
                    LEFT JOIN TBL_RE_DP_RegionalRelatorioBernardo AS R ON B.FILIAL = R.FILIAL	
                    LEFT JOIN TBL_RE_DP_CanaisRelatorioBernardo AS C ON B.CANAL_FINAL = C.CANAL_FINAL	

                    WHERE --c.canal_rb not in ('outros bri', 'tlv ativo bri', 'tlv receptivo bri', 'tlv outros bri','Anteneiros','tlv bri','bri') and
                    b.grupo_plano not in ('CONECTADO', 'RESIDENCIAL', 'COMPLETO','1P','2P BL','3G','CONTROLE','CONTROLE_BOLETO-M1','OI TV PRE-PAGO','PACOTE_DADOS','TV ANTENEIROS','PÓS OCT','PÓS OCT PLANOS','CONTROLE_VOZ','BL+POS','PRÉ-M1','PÓS TOTAL','PÓS TOTAL PLANOS')
                    AND B.INDBD <> 'CANCELAMENTO' 
                    AND LEFT(DATA,6) = @ANOMES
                    
                    GROUP BY INDBD,	
                        GRUPO_PLANO,
                        left(data,6),
                        CASE WHEN ((B.COD_SAP = '1042793') AND (GRUPO_PLANO = 'CONTROLE_BOLETO')) THEN 'Smart Message'
                        ELSE C.CANAL_RB END,
                        R.REGIONAL_AGRUPADA,
                    
                        LTRIM(RTRIM(B.FILIAL)),
                        RIGHT(DATA,2)

                    union all

                    SELECT  'REAL' TIPO_INDICADOR,
                            left(data,6) anomes,
                            INDBD,	
                        GRUPO_PLANO,
                        C.CANAL_RB canal,
                        R.REGIONAL_AGRUPADA regional,
                        'EMP CLI' SEGMENTO,
                        LTRIM(RTRIM(B.UF_CLIENTE)) AS UF,
                        RIGHT(DATA,2) AS DIA,
                        SUM(B.QTD) AS valor

                    FROM TBL_RE_BaseResultadoDiario_Empresarial B	
                    LEFT JOIN TBL_RE_DP_RegionalRelatorioBernardo AS R ON B.UF_CLIENTE = R.FILIAL	
                    LEFT JOIN TBL_RE_DP_CanaisRelatorioBernardo_Empresarial AS C ON B.CANAL_FINAL = C.CANAL_FINAL	

                    WHERE --c.canal_rb not in ('outros bri', 'tlv ativo bri', 'tlv receptivo bri', 'tlv outros bri','Anteneiros','tlv bri','bri') and
                    b.grupo_plano not in ('CONECTADO', 'RESIDENCIAL', 'COMPLETO','1P','2P BL','3G','CONTROLE','CONTROLE_BOLETO-M1','OI TV PRE-PAGO','PACOTE_DADOS','TV ANTENEIROS','PÓS OCT','PÓS OCT PLANOS','CONTROLE_VOZ','BL+POS','PRÉ-M1','PÓS TOTAL','PÓS TOTAL PLANOS')
                    AND B.INDBD <> 'CANCELAMENTO' 
                    AND LEFT(DATA,6)  = @ANOMES
                    
                    GROUP BY INDBD,	
                        GRUPO_PLANO,
                        left(data,6),
                        C.CANAL_RB,
                        R.REGIONAL_AGRUPADA,
                    
                        LTRIM(RTRIM(B.UF_CLIENTE)),
                        RIGHT(DATA,2)

                    union all

                    SELECT  'REAL' TIPO_INDICADOR,
                            left(data,6) anomes,
                            INDBD,	
                        GRUPO_PLANO,
                        C.CANAL_RB canal,
                        R.REGIONAL_AGRUPADA regional,
                        'EMP PDV' SEGMENTO,
                        LTRIM(RTRIM(B.UF_CARTEIRA)) AS UF,
                        RIGHT(DATA,2) AS DIA,
                        SUM(B.QTD) AS valor
                    FROM TBL_RE_BaseResultadoDiario_Empresarial B	
                    LEFT JOIN TBL_RE_DP_RegionalRelatorioBernardo AS R ON B.UF_CARTEIRA = R.FILIAL	
                    LEFT JOIN TBL_RE_DP_CanaisRelatorioBernardo_Empresarial AS C ON B.CANAL_FINAL = C.CANAL_FINAL	

                    WHERE --c.canal_rb not in ('outros bri', 'tlv ativo bri', 'tlv receptivo bri', 'tlv outros bri','Anteneiros','tlv bri','bri') and
                    b.grupo_plano not in ('CONECTADO', 'RESIDENCIAL', 'COMPLETO','1P','2P BL','3G','CONTROLE','CONTROLE_BOLETO-M1','OI TV PRE-PAGO','PACOTE_DADOS','TV ANTENEIROS','PÓS OCT','PÓS OCT PLANOS','CONTROLE_VOZ','BL+POS','PRÉ-M1','PÓS TOTAL','PÓS TOTAL PLANOS')
                    AND B.INDBD <> 'CANCELAMENTO' 
                    AND LEFT(DATA,6) = @ANOMES
                    
                    GROUP BY INDBD,	
                        GRUPO_PLANO,
                        left(data,6),
                        C.CANAL_RB,
                        R.REGIONAL_AGRUPADA,
                        LTRIM(RTRIM(B.UF_CARTEIRA)),
                        RIGHT(DATA,2)
                    ) t
                    /*
                    TRUNCATE TABLE TBL_IND_VAR_ACOMPANHAMENTO_RESIDENCIAL_BASE_NOVA

                    DECLARE @DATA_2 AS VARCHAR (6)

                    SET @DATA_2 = (SELECT MAX(DATA) FROM TBL_IND_VAR_ACOMPANHAMENTO_RESIDENCIAL_BASE)

                    insert into TBL_IND_VAR_ACOMPANHAMENTO_RESIDENCIAL_BASE_NOVA
                SELECT 
                    --TIPO_INDICADOR,
                    DATA,
                    INDBD,
                    case 
                                    when GRUPO_PLANO in ('CONTROLE BOLETO','CONTROLE_BOLETO') then 'CONTROLE_BOLETO' 
                                    when GRUPO_PLANO in ('CONTROLE CARTAO','CONTROLE_CARTAO') then 'CONTROLE_CARTAO' 
                                    when GRUPO_PLANO in ('MOVEL','MÓVEL') then 'MOVEL' 
                                    when GRUPO_PLANO in ('OI GALERA PRÉ','PRÉ-PAGO') then 'PRÉ-PAGO' 
                            else GRUPO_PLANO end GRUPO_PLANO,
                    CANAL,
                    REGIONAL,
                    SEGMENTO,
                    GESTAO GESTÃO,
                    UF,
                    DIA,
                    0 AS COMPROMISSO,
                    SUM(CASE WHEN  TIPO_INDICADOR = 'FORECAST' THEN QTD ELSE 0 END) AS FORECAST,
                    SUM(CASE WHEN  TIPO_INDICADOR = 'META' THEN QTD ELSE 0 END) AS META,
                    SUM(CASE WHEN  TIPO_INDICADOR = 'ORÇAMENTO' THEN QTD ELSE 0 END) AS ORCAMENTO,
                    SUM(CASE WHEN  TIPO_INDICADOR = 'REAL' THEN QTD ELSE 0 END) AS REAL,
                    SUM(CASE WHEN  TIPO_INDICADOR = 'TENDÊNCIA' THEN QTD ELSE 0 END) AS TENDENCIA

                FROM TBL_IND_VAR_ACOMPANHAMENTO_RESIDENCIAL_BASE
                /* REMOVENDO AS VISÕES DE FIBRA EM OBRA - SOLICITAÇÃO DO MARIO - 07/05/2019 */
                WHERE GRUPO_PLANO NOT LIKE '%C' AND GRUPO_PLANO NOT LIKE '%EO' AND GRUPO_PLANO NOT LIKE '%SI' AND GRUPO_PLANO NOT LIKE '%EX OBRA'

                GROUP BY 	DATA,
                    INDBD,
                    GRUPO_PLANO,
                    CANAL,
                    REGIONAL,
                    SEGMENTO,
                    GESTAO,
                    UF,
                    DIA

                UNION ALL

                select 
                    --TIPO_INDICADOR,
                    ANOMES DATA,
                    INDBD,
                    case 
                                    when GRUPO_PLANO in ('CONTROLE BOLETO','CONTROLE_BOLETO') then 'CONTROLE_BOLETO' 
                                    when GRUPO_PLANO in ('CONTROLE CARTAO','CONTROLE_CARTAO') then 'CONTROLE_CARTAO' 
                                    when GRUPO_PLANO in ('MOVEL','MÓVEL') then 'MOVEL' 
                                    when GRUPO_PLANO in ('OI GALERA PRÉ','PRÉ-PAGO') then 'PRÉ-PAGO' 
                            else GRUPO_PLANO end GRUPO_PLANO,
                    CANAL,
                    REGIONAL,
                    CASE WHEN UNIDADE_NEGOCIO = 'EMPRESARIAL' THEN 'EMP PDV' ELSE UNIDADE_NEGOCIO END AS SEGMENTO,
                    GESTÃO,
                    UF,
                    DIA,
                    SUM(QTD) AS COMPROMISSO,
                    0 FORECAST,
                    0 META,
                    0 ORCAMENTO,
                    0 REAL,
                    0 TENDENCIA

                from TBL_PC_BASEMETA_RELATORIO_MANOEL

                WHERE TIPO_INDICADOR = 'COMPROMISSO'
                AND ANOMES >= @DATA_2

                GROUP BY 	--TIPO_INDICADOR,
                    ANOMES,
                    INDBD,
                    GRUPO_PLANO,
                    CANAL,
                    REGIONAL,
                    UNIDADE_NEGOCIO,
                    GESTÃO,
                    UF,
                    DIA*/

                    INSERT INTO TBL_PC_TEMPO_PROCEDURES
                    SELECT
                    'SP_PC_CG_IND_Acompanhamento_Diario_Final' AS [PROCEDURE],
                    'FIM' AS INI_FIM,
                    GETDATE() AS DATA_HORA


                    print '--- Commit ---'	
                    PRINT '-----------------------------------------------------------------------------------------'
                    PRINT '-            FIM CARGA SP_PC_CG_IND_Acompanhamento_Diario_Final   MÊS: '+@ANOMES
                    PRINT '-----------------------------------------------------------------------------------------'
                    
                    
                    --COMMIT
                end
				'''

	dados_conexao = (
		"Driver={SQL Server};"
		f"Server={segredos.db_server};"
		f"Database={segredos.db_name};"
		f"UID={segredos.db_user};"
		f"PWD={segredos.db_pass}"
	)
	conexao = pyodbc.connect(dados_conexao)
	print("Conectado ao banco para alterar a procedure - colocar comentários")
	cursor = conexao.cursor()
	cursor.execute(comando_sql)
	conexao.commit()
	conexao.close()
	print('Conexão Fechada')

def menu():
    print('----------------- Menu de Automacao de Atividades -----------------')
    print('')
    print('1) Verificar as datas no MONITOR DE CARGA e Demonstrativo do Gross')
    print('2) Copiar o Demonstrativo do Gross e montar tabela dinâmica')
    print('3) Executar processo da Nova Fibra')
    print('4) Executar procedures para o Legado')
    print('5) Procedures Finais - usar depois de atualizar a tendência manualmente')
    print('6) Procedures Receita Contratada')
    print('')
    selecionada =  input(('Selecione uma das opções acima: #'))
    print(f'A opção selecionda foi: {selecionada}')
    return selecionada



param = datetime.today().strftime('%Y%m')

opcaoSelecionada = menu()
if opcaoSelecionada == '1':
    print('Iniciando a verificação de datas...')
   
    fim = puxa_dts_cargas()
    BOV_1058 = (f'{fim["BOV_1058.TXT"].split(" ")[0]}')
    BOV_1059 = (f'{fim["BOV_1059.TXT"].split(" ")[0]}')
    BOV_1064 = (f'{fim["BOV_1064.TXT"].split(" ")[0]}')
    BOV_1065 = (f'{fim["BOV_1065.TXT"].split(" ")[0]}')
    BOV_1066 = (f'{fim["BOV_1066.TXT"].split(" ")[0]}')
    BOV_1067 = (f'{fim["BOV_1067.TXT"].split(" ")[0]}')
    BOV_6162 = (f'{fim["HADOOP_6162.TXT"].split(" ")[0]}')
    BOV_6163 = (f'{fim["HADOOP_6163.TXT"].split(" ")[0]}')

    data_mod = data_mod_arquivo('Y:\Demonstrativo Gross.xlsb')
    print(f'HOJE               : {hoje}')
    print(f'Demonstrativo Gross: {data_mod}')
    
    print(f'BOV_1058           : {BOV_1058}')
    print(f'BOV_1059           : {BOV_1059}')
    print(f'BOV_1064           : {BOV_1064}')
    print(f'BOV_1065           : {BOV_1065}')
    print(f'BOV_1066           : {BOV_1066}')
    print(f'BOV_1067           : {BOV_1067}')
    print(f'BOV_6162           : {BOV_6162}')
    print(f'BOV_6163           : {BOV_6163}')
elif opcaoSelecionada == '2':
    if data_mod != hoje:
        ConfirmaContinuar = input('Data do arquivo diferente da data de hoje. Deseja continuar? (S/N):')
        if ConfirmaContinuar == 'N' or ConfirmaContinuar == 'n':
            print('SAIR')
        elif  ConfirmaContinuar == 'S' or ConfirmaContinuar == 's':
            copia_arquivo_renomeia()
            monta_tabdin_demonstrativo_gross()
        else:
            print('Opção inválida')
    elif data_mod == hoje:
        copia_arquivo_renomeia()
        monta_tabdin_demonstrativo_gross()


elif opcaoSelecionada == '3':

    print('\x1b[1;33;44m' + 'Alterando a Procedure SP_PC_NOVA_FIBRA - descomentando o bloco de tendência'+ '\x1b[0m')
    tira_comentario_procedure_nova_fibra_sql()
    
    print('\x1b[1;33;44m' + 'Executando a Procedure SP_PC_NOVA_FIBRA'+ '\x1b[0m')
    executa_procedure_sql('SP_PC_NOVA_FIBRA', param)	
    
    print('\x1b[1;33;44m' + 'Alterando a Procedure SP_PC_NOVA_FIBRA - comentando o bloco de tendência'+ '\x1b[0m')
    coloca_comentario_procedure_nova_fibra_sql()
    
    print('\x1b[1;33;44m' + 'Montando Excel com tabela dinamica e salvando na rede'+ '\x1b[0m')
    montaExcelTendVll()
    
    print('\x1b[1;33;44m' + 'Enviando e-mail'+ '\x1b[0m')
    enviaEmaileAnexo()
    
    print('\x1b[1;32;41m' + 'CONCLUIDO' + '\x1b[0m')
elif opcaoSelecionada == '4':
    executa_procedure_sql('SP_PC_Insert_Tendencia_Auto_Fibra',param)
    print('Concluido. Continue o processo no excel para calcular a tendência de VL e VLL da Fibra Legado.')
elif opcaoSelecionada == '5':
    ATIVAR_TEND_TABLEAU_teste_Jan22()
    executa_procedure_sql('SP_PC_BASES_SHAREPOINT',param)
    ATIVAR_TEND_TABLEAU_teste_Jan22_somenteFibra()
elif opcaoSelecionada == '6':
    executa_procedure_sql('SP_PC_Update_Ticket_Fibra_VAREJO_Tendencia_porRegiao', param)
    executa_procedure_sql('SP_PC_Update_Ticket_Fibra_EMPRESARIAL_Tendencia_porRegiao_IndCombo', param)
    executa_procedure_sql('SP_PC_TBL_RE_RELATORIO_RC_V2_TEND', param)
else:
    print('Opção Inválida')

print('FIM')