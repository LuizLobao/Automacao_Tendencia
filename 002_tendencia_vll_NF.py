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
		

		dt6163 = pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[83]/td[8]').text_content()
		dt6162 = pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[82]/td[8]').text_content()
		
		dw6163 = pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[83]/td[5]').text_content()
		dw6162 = pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[82]/td[5]').text_content()

		di6163 = pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[83]/td[7]').text_content()
		di6162 = pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[82]/td[7]').text_content()
		

		navegador.close()

		datas_fim = [dt6163,dt6162]
		datas_down = [dw6163,dw6162]
		datas_ini = [di6163,di6162]

		#print(dt6163)
		
		return datas_fim, datas_down, datas_ini

def executa_procedure_sql():
	

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

def tira_comentario_procedure_nova_fibra_sql():
	comando_sql='''
				
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

def rotina_se_ok ():
	print('\x1b[1;33;44m' + 'Datas dos arquivos 6163 e 6162 iguais a data de Hoje. Continuando o processo de tendência...'+ '\x1b[0m')
	
	# 2) Rodar procedure com codigo "descomentado"
	#print('\x1b[1;33;44m' + 'Alterando a Procedure SP_PC_NOVA_FIBRA - descomentando o bloco de tendência'+ '\x1b[0m')
	#tira_comentario_procedure_nova_fibra_sql()

	# 3) Rodar procedure SP_PC_NOVA_FIBRA - OK
	print('\x1b[1;33;44m' + 'Executando a Procedure SP_PC_NOVA_FIBRA_COM_TENDENCIA'+ '\x1b[0m')
	executa_procedure_sql()	

	# 4) Rodar procedure com codigo comentado - OK
	#print('\x1b[1;33;44m' + 'Alterando a Procedure SP_PC_NOVA_FIBRA - comentando o bloco de tendência'+ '\x1b[0m')
	#coloca_comentario_procedure_nova_fibra_sql()

	# 5) Rodar query e colocar no excel - OK  |  # 6) salvar na rede - OK
	print('\x1b[1;33;44m' + 'Montando Excel com tabela dinamica e salvando na rede'+ '\x1b[0m')
	montaExcelTendVll()

	# 7) mandar por e-mail - OK
	print('\x1b[1;33;44m' + 'Enviando e-mail'+ '\x1b[0m')
	enviaEmaileAnexo()

	print('\x1b[1;32;41m' + 'CONCLUIDO' + '\x1b[0m')

def principal():
	# 1) Verificar se bases rodaram no MONITOR DE CARGA - OK
	print(f'\x1b[1;32;42m' + 'Verificando datas no Monitor de Cargas...'+ '\x1b[0m')
	fim, down, ini = puxa_dts_cargas()
	dia6163 = (fim[0].split(' ')[0]) #pega somente a data do primeiro registro: 6163
	dia6162 = (fim[1].split(' ')[0]) #pega somente a data do segundo registro: 6162

	#dia6162 = '18/10/2022'
	print(f'Data do 6163: {dia6163}. Data do 6162:{dia6162}')

	tentativas = 1
	while (hoje != dia6163) or (hoje != dia6162):
		tempo_espera = 5
		tempo_inicial = 0

		while tempo_inicial < tempo_espera:
			print(f'\x1b[1;32;41m' + 'Datas Diferentes. Esperando {tempo_espera - tempo_inicial} minutos' + '\x1b[0m')
			time.sleep(60) #1minutos
			tempo_inicial = tempo_inicial + 1

		tentativas = tentativas + 1
		print(f'Verificando datas no Monitor de Cargas...Tentativa número: {tentativas}')
		fim, down, ini = puxa_dts_cargas()
		ia6163 = (fim[0].split(' ')[0]) #pega somente a data do primeiro registro: 6163
		dia6162 = (fim[1].split(' ')[0]) #pega somente a data do segundo registro: 6162
		print(f'Data do 6163: {dia6163}. Dia do 6162:{dia6162}')


	if hoje == dia6163 == dia6162:
		print(f'\x1b[1;32;40m' + 'Arquivos disponíveis. Continuando o processo...' + '\x1b[0m')
		rotina_se_ok()
	else:
		print('\x1b[1;32;41m' + 'ESCAPE Datas Diferentes. Continuar esperando' + '\x1b[0m')

#principal()
rotina_se_ok()