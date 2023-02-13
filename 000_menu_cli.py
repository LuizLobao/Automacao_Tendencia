#TODO trocar query digitada dentro da def por um arquivo SQL
#TODO Salvar status de cada etapa. Só rodar a seguinte se a anterior ja rodou
#TODO estudar a possibilidade de passar uma lista de PROCEDURES e rodar em Loop - desta forma realiza 1 unica conexao

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
import subprocess
from tqdm import tqdm

#FIXME caso deixe o programa rodando de um dia para o outro a variavel não atualiza - causando problemas no dia seguinte 
hoje = datetime.today().strftime('%d/%m/%Y')
AAAAMMDD = datetime.today().strftime('%Y%m%d')
AAAA_MM = datetime.today().strftime('%Y-%m')
AAAAMM = datetime.today().strftime('%Y%m')
resposta = ''

def menu():
	subprocess.run('cls', shell=True)
	print('----------------- Menu de Automacao de Atividades -----------------')
	print('')
	print('1) Verificar as datas no MONITOR DE CARGA e Demonstrativo do Gross')
	print('2) Copiar o Demonstrativo do Gross e montar tabela dinâmica')
	print('3) Executar processo da Nova Fibra')
	print('4) Executar procedures para o Legado')
	print('5) Update Tendências = Real')
	print('6) Procedures Finais - usar depois de atualizar a tendência manualmente')
	print('7) Procedures Receita Contratada')
	print('8) Sair')
	print('-------------------------------------------------------------------')
	selecionada =  input(('Selecione uma das opções acima: #'))
	print(f'A opção selecionda foi: {selecionada}')
	return selecionada

def data_mod_arquivo():
	arquivo1 = f'Demonstrativo Gross_Analitico_{AAAAMM}.csv'
	arquivo = (f'Y:\{arquivo1}')
	modificado = time.strftime('%d/%m/%Y', time.gmtime(os.path.getmtime(arquivo)))
	return (modificado)

def puxa_dts_cargas(em_loop):
	dicio = {"arquivo":"data","HOJE":hoje}
	with sync_playwright() as p:

		navegador = p.chromium.launch(headless=True)
		pagina = navegador.new_page()
		pagina.goto("http://10.20.83.116/aplicacao/monitor/")
		linha = 1
		with tqdm(total=93) as barra_progresso:
			while linha <= 93:
					arquivo=(pagina.locator(f'xpath = //*[@id="mytable"]/tbody/tr[{linha}]/td[9]').text_content())
					DataFim=(pagina.locator(f'xpath = //*[@id="mytable"]/tbody/tr[{linha}]/td[8]').text_content())
					Projeto=(pagina.locator(f'xpath = //*[@id="mytable"]/tbody/tr[{linha}]/td[1]').text_content())
					Status= (pagina.locator(f'xpath = //*[@id="mytable"]/tbody/tr[{linha}]/td[6]').text_content())
					
					if Projeto == 'BASE_FIBRA' or Projeto == 'NOVA_FIBRA':
						if Status == 'Carga em andamento' or Status == 'Carga realizada':
							if DataFim == '\xa0':
								DataFim = '01/01/1900 00:00:00'
								#Status=(pagina.locator(f'xpath = //*[@id="mytable"]/tbody/tr[{linha}]/td[6]').text_content())
							dicio.update({arquivo:DataFim})
					linha += 1
					barra_progresso.update(1)
		navegador.close()
	

	print(dicio)

	dataDemostrativoGross = data_mod_arquivo()

	
	BOV_1067 = (f'{dicio["BOV_1067.TXT"].split(" ")[0]}')
	BOV_1058 = (f'{dicio["BOV_1058.TXT"].split(" ")[0]}')
	BOV_1059 = (f'{dicio["BOV_1059.TXT"].split(" ")[0]}')
	BOV_1065 = (f'{dicio["BOV_1065.TXT"].split(" ")[0]}')
	BOV_1064 = (f'{dicio["BOV_1064.TXT"].split(" ")[0]}')
	BOV_6162 = (f'{dicio["HADOOP_6162.TXT"].split(" ")[0]}')
	BOV_6163 = (f'{dicio["HADOOP_6163.TXT"].split(" ")[0]}')

	print(f"1067: {BOV_1067}")
	print(f"1058: {BOV_1058}")
	print(f"1059: {BOV_1059}")
	print(f"1065: {BOV_1065}")
	print(f"1064: {BOV_1064}")
	print(f"6162: {BOV_6162}")
	print(f"6163: {BOV_6163}")
	print(f'Demonstrativo Gross: {dataDemostrativoGross}')

	if hoje == BOV_1067 == BOV_1058 == BOV_1059 == BOV_1065 == BOV_1064 == BOV_6162 == BOV_6163:
		print('Todos os arquivos da BOV têm a data de hoje...podemos continuar')
	else:
		print('Um ou mais arquivos do BOV NÃO têm a data de hoje...aguardar')
		colocar_puxa_dts_carga_em_loop(em_loop)
			
def colocar_puxa_dts_carga_em_loop(em_loop):
	if em_loop == 'n' or em_loop == 'N':
		resposta = input('Gostaria de colocar o check em Loop ? (s/n):')
		if resposta == 's' or resposta == 'S' or em_loop == 's' or em_loop == 'S':
			print('Esperando 300 segundos = 5 min')
			time.sleep(60) #60 segundos
			print('Esperando 240 segundos = 4 min')
			time.sleep(60) #60 segundos
			print('Esperando 180 segundos = 3 min')
			time.sleep(60) #60 segundos
			print('Esperando 120 segundos = 2 min')
			time.sleep(60) #60 segundos
			print('Esperando 60 segundos = 1 min')
			time.sleep(60) #60 segundos
			puxa_dts_cargas('s')
	if em_loop == 's' or em_loop == 'S':
		print('Esperando 300 segundos = 5 min')
		time.sleep(60) #60 segundos
		print('Esperando 240 segundos = 4 min')
		time.sleep(60) #60 segundos
		print('Esperando 180 segundos = 3 min')
		time.sleep(60) #60 segundos
		print('Esperando 120 segundos = 2 min')
		time.sleep(60) #60 segundos
		print('Esperando 60 segundos = 1 min')
		time.sleep(60) #60 segundos
		puxa_dts_cargas('s')

def copia_arquivo_renomeia():
    shutil.copy(rf"Y:\\Demonstrativo Gross_Analitico_{AAAAMM}.csv", fr'S:\\Resultados\\01_Relatorio Diario\\1 - Base Eventos\\02 - TENDÊNCIA\\Insumos_Tendência\\Demonstrativo Gross_Analitico_{AAAAMMDD}.csv')
    
def monta_tabdin_demonstrativo_gross():
    in_file = f'S:\\Resultados\\01_Relatorio Diario\\1 - Base Eventos\\02 - TENDÊNCIA\\Insumos_Tendência\\Demonstrativo Gross_Analitico_{AAAAMMDD}.csv'
    dest_filename = f'S:\\Resultados\\01_Relatorio Diario\\1 - Base Eventos\\02 - TENDÊNCIA\\Insumos_Tendência\\Demonstrativo Gross_{AAAAMMDD}.xlsx'

    df = pd.read_csv(in_file, decimal=',',sep=';', quotechar='"')
    df1=df[df['TIPO'].str.contains("INSTALACAO")]
    pt_instalacao = df1.query('TIPO == "INSTALACAO"'and 'MERCADO in ("EMPRESARIAL", "VAREJO")').pivot_table(
                                                                                                        values="PROJ", 
                                                                                                        index=["UF"], 
                                                                                                        columns="MERCADO", 
                                                                                                        aggfunc=sum,
                                                                                                        fill_value=0,
                                                                                                        margins=True, margins_name="INSTALACAO",
                                                                                                        )
    df2=df[df['TIPO'].str.contains("MIGRACAO")]
    pt_migracao = df2.query('TIPO == "MIGRACAO"'and 'MERCADO in ("EMPRESARIAL","VAREJO")').pivot_table(
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

def executa_procedure_sql_simples():
	
	dados_conexao = (
		"Driver={SQL Server};"
		f"Server={segredos.db_server};"
		f"Database={segredos.db_name};"
		f"UID={segredos.db_user};"
		f"PWD={segredos.db_pass}"
	)
	conexao = pyodbc.connect(dados_conexao)
	print("Conectado ao banco para executar PROCEDURE")

	cursor = conexao.cursor()
	
	#executar procedure
	procedure = 'SP_PC_NOVA_FIBRA_COM_TENDENCIA'
	dh_inicio_proc = datetime.today().strftime('%Y%m%d %H:%M:%S')
	print(f'Hora inicio execução procedure: {dh_inicio_proc}')
	cursor.execute('SET NOCOUNT ON; EXEC SP_PC_NOVA_FIBRA_COM_TENDENCIA')
	conexao.commit()

	############# LOOP PARA VERIFICAR FIM DA PROCEDURE #############
	def verifica_fim_procedure (procedure, dh_inicio_proc):
		tentativas = 1
		dh_fim_proc = dh_inicio_proc

		dados_conexao = (
			"Driver={SQL Server};"
			f"Server={segredos.db_server};"
			f"Database={segredos.db_name};"
			f"UID={segredos.db_user};"
			f"PWD={segredos.db_pass}"
		)
		conn = pyodbc.connect(dados_conexao)
		cursor = conn.cursor()
		comando_sql = f"SELECT max(DATA_HORA) as DATA_HORA FROM TBL_PC_TEMPO_PROCEDURES WHERE [PROCEDURE] = '{procedure}' AND INI_FIM = 'FIM'"
		
		while dh_fim_proc <= dh_inicio_proc:
	
			print(f'Aguardando Fim da procedure. Tentativa: {tentativas}')
			cursor.execute(comando_sql)
			row = cursor.fetchone()
			dh_fim_proc = row.DATA_HORA.strftime('%Y%m%d %H:%M:%S')

			print(f'Hora fim da execução da procedure: {dh_fim_proc}')

			if (dh_fim_proc > dh_inicio_proc):
				print('Procedure concluida. Continuando...')
				conn.close()
				return
			
			print('Procedure não concluida. Esperar mais 1 min')
			time.sleep(60)
			tentativas = tentativas + 1
		return
	
	
	verifica_fim_procedure(procedure, dh_inicio_proc)
	############# FIM DO LOOP PARA VERIFICAR FIM DA PROCEDURE #############
	
	conexao.close()
	print('Conexão da PROCEDURE Fechada')

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
                            A.SEGMENTO,
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
                                   -- SELECT 
                                   --     * 
                                   -- FROM TBL_IND_VAR_DU_SUMARIZADO
                                   --where ANOMES = @ANOMES
								   /* --------------- ALTERADO EM 11/01/2023 -----------------------*/
								   SELECT
										INDBD,
										ANOMES,
										PRODUTO AS PRODUTO_NOVO,
										SEGMENTO,
										SUM(DU) AS VALOR
								   FROM TBL_PC_DU_PRODUTO_SEGMENTO
								   where ANOMES = @ANOMES
								   GROUP BY INDBD,ANOMES, PRODUTO,SEGMENTO

                            ) DU_MES ON A.INDBD = DU_MES.INDBD
                                        --a.GRUPO_PLANO = DU_MES.PRODUTO_NOVO AND
                                        --case when A.grupo_plano = 'NOVA FIBRA' then 'FIBRA' ELSE A.GRUPO_PLANO END = DU_MES.PRODUTO_NOVO AND
                                        AND A.grupo_plano = DU_MES.PRODUTO_NOVO
										--AND CASE WHEN A.CANAL_RB = 'Smart Message' THEN 'TLV Outros' ELSE A.CANAL_RB END = DU_MES.CANAL 
										--AND A.REGIONAL = DU_MES.REGIONAL 
										AND A.DATA = DU_MES.ANOMES
										AND CASE WHEN A.SEGMENTO IN ('EMP CLI','EMP PDV') THEN 'EMPRESARIAL' ELSE A.SEGMENTO END = DU_MES.SEGMENTO

                    /* DU POR DIA PARA REGIONAL - CANAL */
                    left JOIN (/*SELECT ANOMESDIA,
                                    left(ANOMESDIA,6) ANOMES,
                                    INDBD,
                                    CASE WHEN DU_DIA.PRODUTO IN ('OI GALERA PRÉ', 'PRÉ-PAGO') THEN 'PRÉ-PAGO' ELSE DU_DIA.PRODUTO END PRODUTO_NOVO,
                                    CANAL,
                                    REGIONAL,
                                    valor
                                FROM TBL_pc_du AS DU_DIA
                                            where left(ANOMESDIA,6) = @ANOMES*/
								SELECT
										ANOMESDIA,
										INDBD,
										ANOMES,
										PRODUTO AS PRODUTO_NOVO,
										SEGMENTO,
										SUM(DU) AS VALOR
								   FROM TBL_PC_DU_PRODUTO_SEGMENTO
								   where ANOMES = @ANOMES
								   GROUP BY ANOMESDIA,INDBD,ANOMES, PRODUTO,SEGMENTO

                                ) C ON 
                                        A.INDBD = C.INDBD AND 
                                        A.grupo_plano = C.PRODUTO_NOVO AND
										--case when A.grupo_plano = 'NOVA FIBRA' then 'FIBRA' ELSE A.GRUPO_PLANO END = C.PRODUTO_NOVO AND
                                        --CASE WHEN A.CANAL_RB = 'Smart Message' THEN 'TLV Outros' ELSE A.CANAL_RB END = C.CANAL AND
                                        --A.REGIONAL = C.REGIONAL and
                                        CASE WHEN A.SEGMENTO IN ('EMP CLI','EMP PDV') THEN 'EMPRESARIAL' ELSE A.SEGMENTO END = C.SEGMENTO and
										A.DATA = C.ANOMES
                    /* TOTAL DU POR PRODUTOS */
                    left JOIN ( SELECT 
                                            ANOMES,
                                            INDBD,
                                            PRODUTO,
											SEGMENTO,
                                            SUM(DU) VALOR
                                    FROM TBL_PC_DU_PRODUTO_SEGMENTO
                                    where ANOMES = @ANOMES
                                    group by 
                                            ANOMES,
                                            INDBD,
                                            PRODUTO,
											SEGMENTO
								  /*  SELECT 
                                            LEFT(ANOMESDIA,6) ANOMES,
                                            INDBD,
                                            PRODUTO_NOVO,
                                            SUM(cast(rtrim(ltrim(replace(VALOR,',','.'))) as float)) VALOR 
                                    FROM TBL_PC_DU_PRODUTO
                                    where left(ANOMESDIA,6) = @ANOMES
                                    group by 
                                            LEFT(ANOMESDIA,6),
                                            INDBD,
                                            PRODUTO_NOVO*/

                            ) DU_MES_PRODUTO ON	A.INDBD = DU_MES_PRODUTO.INDBD AND 
                                                    A.grupo_plano = DU_MES_PRODUTO.PRODUTO AND
                                                    A.DATA = DU_MES_PRODUTO.ANOMES AND
													 CASE WHEN A.SEGMENTO IN ('EMP CLI','EMP PDV') THEN 'EMPRESARIAL' ELSE A.SEGMENTO END = DU_MES_PRODUTO.SEGMENTO
													-- AND
													--A.SEGMENTO = DU_MES_PRODUTO.SEGMENTO

                    /* DU POR DIA PARA PRODUTO */
                    left JOIN ( SELECT 
                                            ANOMESDIA,
											ANOMES,
                                            INDBD,
                                            PRODUTO as PRODUTO_NOVO,
											SEGMENTO,
                                            SUM(DU) VALOR
                                    FROM TBL_PC_DU_PRODUTO_SEGMENTO
                                    where ANOMES = @ANOMES
                                    group by 
                                            ANOMESDIA,
											ANOMES,
                                            INDBD,
                                            PRODUTO,
											SEGMENTO
								  /*  SELECT 
                                            LEFT(ANOMESDIA,6) ANOMES,
                                            INDBD,
                                            PRODUTO_NOVO,
                                            SUM(cast(rtrim(ltrim(replace(VALOR,',','.'))) as float)) VALOR 
                                    FROM TBL_PC_DU_PRODUTO
                                    where left(ANOMESDIA,6) = @ANOMES
                                    group by 
                                            LEFT(ANOMESDIA,6),
                                            INDBD,
                                            PRODUTO_NOVO*/

                            ) P ON	A.INDBD = P.INDBD AND 
                                    A.grupo_plano = P.PRODUTO_NOVO AND
                                    A.DATA = P.ANOMES AND
									 CASE WHEN A.SEGMENTO IN ('EMP CLI','EMP PDV') THEN 'EMPRESARIAL' ELSE A.SEGMENTO END = P.SEGMENTO
									-- AND
									--A.SEGMENTO = DU_MES_PRODUTO.SEGMENTO
                    
                                    
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
                            a.SEGMENTO,
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
                                   -- SELECT 
                                   --     * 
                                   -- FROM TBL_IND_VAR_DU_SUMARIZADO
                                   --where ANOMES = @ANOMES
				/* --------------- ALTERADO EM 11/01/2023 -----------------------*/
				SELECT
					INDBD,
					ANOMES,
					PRODUTO AS PRODUTO_NOVO,
					SEGMENTO,
					SUM(DU) AS VALOR
				FROM TBL_PC_DU_PRODUTO_SEGMENTO
				where ANOMES = @ANOMES
				GROUP BY INDBD,ANOMES, PRODUTO,SEGMENTO

                            ) DU_MES ON A.INDBD = DU_MES.INDBD
                                        --a.GRUPO_PLANO = DU_MES.PRODUTO_NOVO AND
                                        --case when A.grupo_plano = 'NOVA FIBRA' then 'FIBRA' ELSE A.GRUPO_PLANO END = DU_MES.PRODUTO_NOVO AND
                                        AND A.grupo_plano = DU_MES.PRODUTO_NOVO
					--AND CASE WHEN A.CANAL_RB = 'Smart Message' THEN 'TLV Outros' ELSE A.CANAL_RB END = DU_MES.CANAL 
					--AND A.REGIONAL = DU_MES.REGIONAL 
					AND A.DATA = DU_MES.ANOMES
					AND CASE WHEN A.SEGMENTO IN ('EMP CLI','EMP PDV') THEN 'EMPRESARIAL' ELSE A.SEGMENTO END = DU_MES.SEGMENTO

                    /* DU POR DIA PARA REGIONAL - CANAL */
                    left JOIN (/*
                                        SELECT ANOMESDIA,
                                        left(ANOMESDIA,6) ANOMES,
                                        INDBD,
                                        CASE WHEN DU_DIA.PRODUTO IN ('OI GALERA PRÉ', 'PRÉ-PAGO') THEN 'PRÉ-PAGO' ELSE DU_DIA.PRODUTO END PRODUTO_NOVO,
                                        CANAL,
                                        REGIONAL,
                                        valor
                                        FROM TBL_pc_du AS DU_DIA
                                        where left(ANOMESDIA,6) = @ANOMES
                                */
					SELECT
					        ANOMESDIA,
					        INDBD,
					        ANOMES,
					        PRODUTO AS PRODUTO_NOVO,
					        SEGMENTO,
					        SUM(DU) AS VALOR
					FROM TBL_PC_DU_PRODUTO_SEGMENTO
					where ANOMES = @ANOMES
					GROUP BY ANOMESDIA,INDBD,ANOMES, PRODUTO,SEGMENTO

                                ) C ON 
                                        A.INDBD = C.INDBD AND 
                                        A.grupo_plano = C.PRODUTO_NOVO AND
					--case when A.grupo_plano = 'NOVA FIBRA' then 'FIBRA' ELSE A.GRUPO_PLANO END = C.PRODUTO_NOVO AND
                                        --CASE WHEN A.CANAL_RB = 'Smart Message' THEN 'TLV Outros' ELSE A.CANAL_RB END = C.CANAL AND
                                        --A.REGIONAL = C.REGIONAL and
                                        CASE WHEN A.SEGMENTO IN ('EMP CLI','EMP PDV') THEN 'EMPRESARIAL' ELSE A.SEGMENTO END = C.SEGMENTO and
					A.DATA = C.ANOMES
                    
                    /* TOTAL DU POR PRODUTOS */
                    left JOIN (SELECT 
                                ANOMES,
                                INDBD,
                                PRODUTO,
				SEGMENTO,
                                SUM(DU) VALOR
                               FROM TBL_PC_DU_PRODUTO_SEGMENTO
                               where ANOMES = @ANOMES
                               group by 
                                        ANOMES,
                                        INDBD,
                                        PRODUTO,
					SEGMENTO
					/*  SELECT 
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
                                        */

                            ) DU_MES_PRODUTO ON	A.INDBD = DU_MES_PRODUTO.INDBD AND 
                                                    A.grupo_plano = DU_MES_PRODUTO.PRODUTO AND
                                                    A.DATA = DU_MES_PRODUTO.ANOMES AND
						 CASE WHEN A.SEGMENTO IN ('EMP CLI','EMP PDV') THEN 'EMPRESARIAL' ELSE A.SEGMENTO END = DU_MES_PRODUTO.SEGMENTO
						-- AND
						--A.SEGMENTO = DU_MES_PRODUTO.SEGMENTO

                    /* DU POR DIA PARA PRODUTO */
                    left JOIN ( SELECT 
                                            ANOMESDIA,
											ANOMES,
                                            INDBD,
                                            PRODUTO as PRODUTO_NOVO,
											SEGMENTO,
                                            SUM(DU) VALOR
                                    FROM TBL_PC_DU_PRODUTO_SEGMENTO
                                    where ANOMES = @ANOMES
                                    group by 
                                            ANOMESDIA,
											ANOMES,
                                            INDBD,
                                            PRODUTO,
											SEGMENTO
								  /*  SELECT 
                                            LEFT(ANOMESDIA,6) ANOMES,
                                            INDBD,
                                            PRODUTO_NOVO,
                                            SUM(cast(rtrim(ltrim(replace(VALOR,',','.'))) as float)) VALOR 
                                    FROM TBL_PC_DU_PRODUTO
                                    where left(ANOMESDIA,6) = @ANOMES
                                    group by 
                                            LEFT(ANOMESDIA,6),
                                            INDBD,
                                            PRODUTO_NOVO*/

                            ) P ON	A.INDBD = P.INDBD AND 
                                    A.grupo_plano = P.PRODUTO_NOVO AND
                                    A.DATA = P.ANOMES AND
									 CASE WHEN A.SEGMENTO IN ('EMP CLI','EMP PDV') THEN 'EMPRESARIAL' ELSE A.SEGMENTO END = P.SEGMENTO
									-- AND
									--A.SEGMENTO = DU_MES_PRODUTO.SEGMENTO
                    
                                    
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

def atualiza_TB_VALIDA_CARGA_TENDENCIA():
	comando_sql='update TB_VALIDA_CARGA_TENDENCIA set DATA_CARGA = convert(varchar, getdate(), 120 )'
	dados_conexao = (
		"Driver={SQL Server};"
		f"Server={segredos.db_server};"
		f"Database={segredos.db_name};"
		f"UID={segredos.db_user};"
		f"PWD={segredos.db_pass}"
	)
	conexao = pyodbc.connect(dados_conexao)
	print("Conectado ao banco para dar update")
	cursor = conexao.cursor()
	cursor.execute(comando_sql)
	conexao.commit()
	conexao.close()
	print('Conexão Fechada')


param = datetime.today().strftime('%Y%m')
opcaoSelecionada = 0
while opcaoSelecionada != 8:
	opcaoSelecionada = menu()
	if opcaoSelecionada == '1':
		print('Iniciando a verificação de datas...')
		puxa_dts_cargas('n')
		a = input('Tecle qualquer tecla para continuar...')

	elif opcaoSelecionada == '2':
		print('Opção 2...')
		copia_arquivo_renomeia()
		monta_tabdin_demonstrativo_gross()
		a = input('Tecle qualquer tecla para continuar...')

	elif opcaoSelecionada == '3':
		print('Opção 3...')
		executa_procedure_sql_simples()
		#montaExcelTendVll()
		#enviaEmaileAnexo()
		a = input('Tecle qualquer tecla para continuar...')

	elif opcaoSelecionada == '4':
		print('Opção 4...')
		param = datetime.today().strftime('%Y%m')
		executa_procedure_sql('SP_PC_Insert_Tendencia_Auto_Fibra',param)
		a = input('Tecle qualquer tecla para continuar...')

	elif opcaoSelecionada == '5':
		print('Opção 5...')
		executa_procedure_sql('SP_PC_TEND_IGUAL_REAL_FIBRA_EMPRESARIAL',param)
		executa_procedure_sql('SP_PC_TEND_IGUAL_REAL_FIBRA_VAREJO',param)
		executa_procedure_sql('SP_PC_TEND_IGUAL_REAL_NOVA_FIBRA',param)
		executa_procedure_sql('SP_PC_TEND_IGUAL_REAL_TABELAS_FIBRA',param)
		a = input('Tecle qualquer tecla para continuar...')

	elif opcaoSelecionada == '6':
		print('Opção 6...')
		ATIVAR_TEND_TABLEAU_teste_Jan22()
		executa_procedure_sql('SP_PC_BASES_SHAREPOINT',param)
		ATIVAR_TEND_TABLEAU_teste_Jan22_somenteFibra()
		atualiza_TB_VALIDA_CARGA_TENDENCIA()
		a = input('Tecle qualquer tecla para continuar...')

	elif opcaoSelecionada == '7':
		print('Opção 7...')
		proc = 'SP_PC_Update_Ticket_Fibra_VAREJO_Tendencia_porRegiao'
		executa_procedure_sql(proc, param)
		proc = 'SP_PC_Update_Ticket_Fibra_EMPRESARIAL_Tendencia_porRegiao_IndCombo'
		executa_procedure_sql(proc, param)

		proc = 'SP_PC_Update_Ticket_Fibra_VAREJO_DIARIO_porRegiao'
		executa_procedure_sql(proc, param)

		proc = 'SP_PC_Update_Ticket_Fibra_EMPRESARIAL_DIARIO_porRegiao_IndCombo'
		executa_procedure_sql(proc, param)

		proc = 'SP_PC_TBL_RE_RELATORIO_RC_V2_TEND'
		executa_procedure_sql(proc, param)
		a = input('Tecle qualquer tecla para continuar...')

	elif opcaoSelecionada == '8':
		print('Opção 8...')
		break
	else:
		print('Opção Inválida')

print('FIM')
