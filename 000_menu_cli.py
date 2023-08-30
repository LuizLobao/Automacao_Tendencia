import openpyxl
import os
import pandas as pd
import pyodbc
import requests
import segredos
import shutil
import subprocess
import time
import win32com.client as win32

from datetime import date, datetime, timedelta
from PIL import ImageGrab
from playwright.sync_api import sync_playwright
from tqdm import tqdm
from urllib.parse import quote

num= segredos.num
key = segredos.key

#FIXME caso deixe o programa rodando de um dia para o outro a variavel não atualiza - causando problemas no dia seguinte 
hoje = (datetime.today()- timedelta(days=0)).strftime('%d/%m/%Y') 
AAAAMMDD = (datetime.today()- timedelta(days=0)).strftime('%Y%m%d') 
AAAA_MM = (datetime.today()- timedelta(days=0)).strftime('%Y-%m') 
AAAAMM = (datetime.today()- timedelta(days=0)).strftime('%Y%m')
resposta = ''

def atualiza_variaveis_data():
	hoje = (datetime.today()- timedelta(days=0)).strftime('%d/%m/%Y') 
	AAAAMMDD = (datetime.today()- timedelta(days=0)).strftime('%Y%m%d') 
	AAAA_MM = (datetime.today()- timedelta(days=0)).strftime('%Y-%m') 
	AAAAMM = (datetime.today()- timedelta(days=0)).strftime('%Y%m')

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

def menu():
	subprocess.run('cls', shell=True)
	print('----------------- Menu de Automacao de Atividades -----------------')
	print('')
	print(f'Data de Hoje: {hoje}')
	print('')
	print('1) Verificar as datas no MONITOR DE CARGA e Demonstrativo do Gross')
	print('2) Verifica datas para o JETL - Processo CDO')
	print('3) Copiar o Demonstrativo do Gross e montar tabela dinâmica')
	print('4) Executar processo da Nova Fibra - rodar depois de executar o JETL')
	print('5) Processos CDO - rodar depois de executar o JETL e depois da etapa 4')
	print('6) Executar procedures para o Legado')
	print('7) Update Tendências = Real')
	print('---------------------------------------------------------------------')
	print('      Rodar o processo de tendencia no excel e UPDATES               ')
	print('---------------------------------------------------------------------')
	print('8) Procedures Finais - usar depois de atualizar a tendência manualmente')
	print('9) Procedures Receita Contratada')
	print('10) Envia E-mail Tendencias Liberadas')
	print('11) Envia Lista de PDV Outros')
	print('12) Gravar Tendência CDO no Histórico')
	print('13) Sair')
	print('-------------------------------------------------------------------')
	selecionada =  input(('Selecione uma das opções acima: #'))
	print(f'A opção selecionda foi: {selecionada}')
	return selecionada

def data_mod_arquivo():
    arquivo1 = f'Demonstrativo Gross_Analitico_{AAAAMM}.csv'
    arquivo = f'Y:\{arquivo1}'

    # Verifica o tamanho do arquivo em KB
    tamanho_kb = os.path.getsize(arquivo) / 1024

    if tamanho_kb <= 10:
        print('Erro - arquivo tem tamanho menor ou igual a 10 KB')
        return None
    else:
        try:
            modificado = time.strftime('%d/%m/%Y', time.gmtime(os.path.getmtime(arquivo)))
            return modificado
        except:
            print('Erro - arquivo Demonstrativo Gross nao encontrado')
            data_erro = '01/01/1900 00:00:00'
            return data_erro

def puxa_dts_jetl(relatorio, aplicacao):
    with sync_playwright() as p:
        navegador = p.chromium.launch(headless=True)
        pagina = navegador.new_page()
        pagina.goto(f"https://portalbi.telemar/AdminRelBatchCadastroJETL.aspx?idRelatorio={relatorio}&idAplicacao={aplicacao}")

        pagina.wait_for_load_state('load')

        linha = 2
        while True:  # Run until the end of the page (infinite loop)
            try:
                # Try to locate the elements for the current row
                # If the elements are not found, it will raise an exception, and the loop will stop
                # Otherwise, it will continue to the next row
                periodo_de = pagina.locator(f'xpath = //*[@id="lbl1"]/table/tbody/tr[{linha}]/td[5]/font/b').text_content()
                periodo_ate = pagina.locator(f'xpath = //*[@id="lbl1"]/table/tbody/tr[{linha}]/td[6]/font/b').text_content()
                data_criacao = pagina.locator(f'xpath = //*[@id="lbl1"]/table/tbody/tr[{linha}]/td[8]/font/b').text_content()

                try:
                    imagem = pagina.query_selector(f'//*[@id="lbl1"]/table/tbody/tr[{linha}]/td[10]/font/b/img')
                    texto_alt = imagem.get_attribute('src')
                except:
                    imagem = pagina.query_selector(f'//*[@id="lbl1"]/table/tbody/tr[{linha}]/td[10]/font/b/a/img')
                    texto_alt = imagem.get_attribute('src')

                print(relatorio, periodo_de, periodo_ate, data_criacao, texto_alt)
                linha += 1

            except Exception as e:
                # If an exception is raised, it means we have reached the end of the page, so break the loop.
                print("Fim da página.")
                break

        print('')
        navegador.close()

def puxa_dts_cargas(em_loop):
	
	try:
		dicio = {"arquivo":"data","HOJE":hoje,"BOV_1066.TXT":'01/01/1900 00:00:00',"BOV_1065.TXT":'01/01/1900 00:00:00',"BOV_1059.TXT":'01/01/1900 00:00:00',"BOV_1067.TXT":'01/01/1900 00:00:00',"BOV_1064.TXT":'01/01/1900 00:00:00',"BOV_1058.TXT":'01/01/1900 00:00:00',"HADOOP_6162.TXT":'01/01/1900 00:00:00',"HADOOP_6163.TXT":'01/01/1900 00:00:00'}
		with sync_playwright() as p:

			navegador = p.chromium.launch(headless=True)
			pagina = navegador.new_page()
			pagina.goto("http://10.25.114.33/aplicacao/monitor/")
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


		#print(dicio)
		dataDemostrativoGross = data_mod_arquivo()
		BOV_1067 = (f'{dicio["BOV_1067.TXT"].split(" ")[0]}')
		BOV_1058 = (f'{dicio["BOV_1058.TXT"].split(" ")[0]}')
		BOV_1059 = (f'{dicio["BOV_1059.TXT"].split(" ")[0]}')
		BOV_1065 = (f'{dicio["BOV_1065.TXT"].split(" ")[0]}')
		BOV_1064 = (f'{dicio["BOV_1064.TXT"].split(" ")[0]}')
		BOV_6162 = (f'{dicio["HADOOP_6162.TXT"].split(" ")[0]}')
		BOV_6163 = (f'{dicio["HADOOP_6163.TXT"].split(" ")[0]}')

		if BOV_1067 != hoje:
			print('\x1b[1;33;41m' + f"1067: {BOV_1067}" + '\x1b[0m')
		else:
			print('\x1b[1;32;40m' + f"1067: {BOV_1067}" + '\x1b[0m')

		if BOV_1058 != hoje:
			print('\x1b[1;33;41m' + f"1058: {BOV_1058}" + '\x1b[0m')
		else:
			print('\x1b[1;32;40m' + f"1058: {BOV_1058}" + '\x1b[0m')	

		if BOV_1059 != hoje:
			print('\x1b[1;33;41m' + f"1059: {BOV_1059}" + '\x1b[0m')
		else:
			print('\x1b[1;32;40m' + f"1059: {BOV_1059}" + '\x1b[0m')

		if BOV_1065 != hoje:
			print('\x1b[1;33;41m' + f"1065: {BOV_1065}" + '\x1b[0m')
		else:
			print('\x1b[1;32;40m' + f"1065: {BOV_1065}" + '\x1b[0m')

		if BOV_1064 != hoje:
			print('\x1b[1;33;41m' + f"1064: {BOV_1064}" + '\x1b[0m')
		else:
			print('\x1b[1;32;40m' + f"1064: {BOV_1064}" + '\x1b[0m')

		if BOV_6162 != hoje:
			print('\x1b[1;33;41m' + f"6162: {BOV_6162}" + '\x1b[0m')
		else:
			print('\x1b[1;32;40m' + f"6162: {BOV_6162}" + '\x1b[0m')

		if BOV_6163 != hoje :
			print('\x1b[1;33;41m' + f"6163: {BOV_6163}"+ '\x1b[0m')
		else:
			print('\x1b[1;32;40m' + f"6163: {BOV_6163}"+ '\x1b[0m')

		if dataDemostrativoGross != hoje :
			print('\x1b[1;33;41m' + f"Demonstrativo Gross: {dataDemostrativoGross}"+ '\x1b[0m')
		else:
			print('\x1b[1;32;40m' + f"Demonstrativo Gross: {dataDemostrativoGross}"+ '\x1b[0m')


		if hoje == BOV_1067 == BOV_1058 == BOV_1059 == BOV_1065 == BOV_1064 == BOV_6162 == BOV_6163:
			msg = f'Todos os arquivos da BOV têm a data de hoje...podemos continuar : {datetime.today()}'
			print(msg)
			enviaWhats(msg, num, key)
		else:
			print('Um ou mais arquivos do BOV NÃO têm a data de hoje...aguardar')
			colocar_puxa_dts_carga_em_loop(em_loop)
	except:
		print('ERRO COM O MONITOR DE CARGA')
	
def colocar_puxa_dts_carga_em_loop(em_loop):
	if em_loop.lower() == 'n':
		resposta = input('Gostaria de colocar o check em Loop ? (s/n):')
		if resposta.lower() == 's' or em_loop.lower() == 's':
			for t in [300, 240, 180, 120, 60]:
				print(f'Esperando {t} segundos = {t//60} min')
				time.sleep(60)
			puxa_dts_cargas('s')
	if em_loop.lower() == 's':
		#continua = input('Continuar no loop (S/N):')
		#if continua.lower() == 's':
		for t in [300, 240, 180, 120, 60]:
			print(f'Esperando {t} segundos = {t//60} min')
			time.sleep(60)
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
                                                                                                        margins=True, 
																										margins_name="INSTALACAO",
                                                                                                        )
    df2=df[df['TIPO'].str.contains("MIGRACAO")]
    pt_migracao = df2.query('TIPO == "MIGRACAO"'and 'MERCADO in ("EMPRESARIAL","VAREJO")').pivot_table(
                                                                                                        values=["PROD","PROJ"], 
                                                                                                        #index=["UF"], 
                                                                                                        index="MERCADO", 
                                                                                                        aggfunc=sum,
                                                                                                        fill_value=0,
                                                                                                        margins=True, 
																										margins_name="MIGRACAO",
                                                                                                        )
    with pd.ExcelWriter(dest_filename) as writer:
        pt_instalacao.to_excel(writer, sheet_name="TabDin",startcol=0, startrow=0)
        pt_migracao.to_excel(writer, sheet_name="TabDin",startcol=6, startrow=0)

def executa_procedure_sql_combinada(nome_procedure, param=None):
    conexao = criar_conexao()
    print('\x1b[1;33;42m' + 'Conexão realizada ao banco de dados' + '\x1b[0m')
    inicio_procedure = datetime.today()
    
    try:
        cursor = conexao.cursor()

        if param:
            # Executa a procedure com parâmetro
            
            print('\x1b[1;33;44m' + f'Executando a Procedure {nome_procedure} para o parâmetro: {param} '+ '\x1b[0m')
            print(f'Iniciando execução em: {inicio_procedure}')
            cursor.execute(f'SET NOCOUNT ON; EXEC {nome_procedure} {param}')
        else:
            # Executa a procedure sem parâmetro
            print(f'Executando a Procedure {nome_procedure}')
            cursor.execute(f'SET NOCOUNT ON; EXEC {nome_procedure}')

        conexao.commit()
        fim_procedure = datetime.today()
        print(f"Procedure executada em {fim_procedure - inicio_procedure} tempo")
    
    except Exception as e:
        print(f"Erro ao executar a procedure: {e}")
    
    finally:
        conexao.close()
        print('\x1b[1;33;41m' + 'Conexão Fechada'+ '\x1b[0m')

def montaExcelTendVll():
	comando_sql = '''SELECT DATA,
					FILIAL,
					SUM(QTD_DEV) AS [TEND]
					FROM TBL_CDO_TEND_NOVA_FIBRA_VLL
					WHERE LEFT(DATA,6) = (SELECT MAX(left(data, 6)) FROM TBL_CDO_TEND_NOVA_FIBRA_VLL)
					GROUP BY DATA, FILIAL'''

	conexao = criar_conexao()


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

def executa_arquivo_sql(arquivo_sql):
    conexao = criar_conexao()
    print("Conectado ao banco para alterar a procedure - retirar comentários")

    try:
        # Defina o caminho do arquivo .sql
        caminho_arquivo_sql = arquivo_sql

        with open(caminho_arquivo_sql, 'r', encoding='utf-8') as arquivo:
            conteudo_sql = arquivo.read()
            cursor = conexao.cursor()
            cursor.execute(conteudo_sql)
            conexao.commit()

    except Exception as e:
        print(f"Erro ao executar o arquivo SQL: {e}")
    
    finally:
        conexao.close()
        print('Conexão Fechada')

def atualiza_TB_VALIDA_CARGA_TENDENCIA():
	comando_sql='update TB_VALIDA_CARGA_TENDENCIA set DATA_CARGA = convert(varchar, getdate(), 120 )'

	conexao = criar_conexao()
	print("Conectado ao banco para dar update")
	cursor = conexao.cursor()
	cursor.execute(comando_sql)
	conexao.commit()
	conexao.close()
	print('Conexão Fechada')

def enviaEmailFimProcesso():		
	workbook_path = r'S:\Resultados\01_Relatorio Diario\1 - Base Eventos\02 - TENDÊNCIA\TEND_FIBRA_e_UPDATEs.xlsx'
	attachment_path = r'S:\Resultados\01_Relatorio Diario\1 - Base Eventos\02 - TENDÊNCIA\tend_email.xlsx'
	image_path = r'C:\Users\oi066724\Documents\Python\Automacao_Tendencia\paste.png'

	excel = win32.Dispatch('Excel.Application')

	wb = excel.Workbooks.Open(workbook_path)
	sheet = wb.Sheets.Item(3)
	excel.visible = 1

	copyrange = sheet.Range('A23:D45')

	copyrange.CopyPicture(Appearance=1, Format=2)
	ImageGrab.grabclipboard().save('paste.png')
	excel.Quit()

	# create a html body template and set the **src** property with `{}` so that we can use
	# python string formatting to insert a variable
	html_body = """
		<div>
			Tendências Liberadas no banco de dados BDIntelicanais (Infocana).
		</div>
		<br>
		<br>
		<div>
			<img src={}></img>
		</div>
	"""

	# startup and instance of outlook
	outlook = win32.Dispatch('Outlook.Application')

	# create a message
	message = outlook.CreateItem(0)

	# set the message properties
	message.To = segredos.lista_email_fim_processo
	message.Subject = f'Tendências Liberadas: {hoje}!'
	message.HTMLBody = html_body.format(image_path)
	message.Attachments.Add(attachment_path)

	# display the message to review
	message.Display()

	# save or send the message
	message.Send()
	print("Email Enviado")

def copiar_colar_excel(origem, destino, planilha_origem):
    # Abrir arquivo de origem
    wb_origem = openpyxl.load_workbook(origem, data_only=True)
    ws_origem = wb_origem[planilha_origem]

    # Selecionar o range de células a ser copiado
    range_copia = ws_origem['A23':'D49']

    # Abrir arquivo de destino
    wb_destino = openpyxl.Workbook()
    ws_destino = wb_destino.active

    # Copiar e colar como valor
    for row in range_copia:
        nova_linha = []
        for cell in row:
            nova_linha.append(cell.value)
        ws_destino.append(nova_linha)

    # Salvar arquivo de destino
    wb_destino.save(destino)

def enviaWhats (mensagem, numero, apikey):
    requests.get(
    url=f'https://api.callmebot.com/whatsapp.php?phone={numero}&text={quote(mensagem)}&apikey={apikey}'
    )

def EnviaPDVOutros():
	comando_sql = 'select * from [VW_COD_SAP_OUTROS] order by qtd desc'

	conexao = criar_conexao()
	cursor = conexao.cursor()
	df=pd.read_sql(comando_sql, conexao)
	os.makedirs('PDV_OUTROS', exist_ok=True)  
	df.to_csv('PDV_OUTROS/pdv_outros.csv', sep=';', decimal=',') 
	attachment_path = r'C:\Users\oi066724\Documents\Python\Automacao_Tendencia\PDV_OUTROS\pdv_outros.csv'

	# create a html body template and set the **src** property with `{}` so that we can use
	# python string formatting to insert a variable
	html_body = """
		<div>
			Caros, Segue a lista de PDVs que estão aparecendo como OUTROS na BOV.
			Favor verificar e atualizar a classificação dos mesmos.
		</div>
		<br>
		<br>
	"""

	# startup and instance of outlook
	outlook = win32.Dispatch('Outlook.Application')

	# create a message
	message = outlook.CreateItem(0)

	# set the message properties
	message.To = segredos.lista_email_pdv_outros
	message.Subject = f'PDVs Outros: {hoje}!'
	message.HTMLBody = html_body.format()
	message.Attachments.Add(attachment_path)

	# display the message to review
	message.Display()

	# save or send the message
	message.Send()

param = AAAAMM
opcaoSelecionada = 0
while opcaoSelecionada != 13:
	opcaoSelecionada = menu()
	if opcaoSelecionada == '1':
		print('Iniciando a verificação de datas...')
		puxa_dts_cargas('n')
		a = input('Tecle qualquer tecla para continuar...')

	elif opcaoSelecionada == '2':
		print('Opção 2...Verificando datas para o JETL')
		print('Rel  DE        ATE       CRIACAO              IMAGEM')
		puxa_dts_jetl(6163,11)
		puxa_dts_jetl(6162,11)
		puxa_dts_jetl(1059,1)
		puxa_dts_jetl(1065,1)
		print('Caso todos os arquivos com a data de HOJE estiverem com a imagem DOWNLOAD, iniciar o JOB no JETL')
		a = input('Tecle qualquer tecla para continuar...')

	elif opcaoSelecionada == '3':
		print('Opção 3...')
		copia_arquivo_renomeia()
		monta_tabdin_demonstrativo_gross()
		a = input('Tecle qualquer tecla para continuar...')

	elif opcaoSelecionada == '4':
		print('Opção 4...')
		print('Executando procedure SP_CDO_TEND_VL_VLL_FIBRA_NOVA_FIBRA...')
		executa_procedure_sql_combinada('SP_CDO_TEND_VL_VLL_FIBRA_NOVA_FIBRA')
		
		print('Montando Excel da TENDENCIA de VLL...')
		montaExcelTendVll()
		
		print('Enviando e-mail para Amado ...')
		enviaEmaileAnexo()
		
		a = input('Tecle qualquer tecla para continuar...')
	
	elif opcaoSelecionada == '5':
		print('Opção 5...Processos CDO')
		print('TENDENCIA = REAL para tabela CDO')
		executa_procedure_sql_combinada('SP_CDO_0001_TENDENCIA_IGUAL_REAL',param)
		print('Puxa TENDENCIA VLL NOVA FIBRA para tabela CDO')
		executa_procedure_sql_combinada('SP_CDO_0002_TENDENCIA_VLL',param)

		a = input('Tecle qualquer tecla para continuar...')

	elif opcaoSelecionada == '6':
		print('Opção 6...')
		param = datetime.today().strftime('%Y%m')
		executa_procedure_sql_combinada('SP_PC_Insert_Tendencia_Auto_Fibra', param)
		a = input('Tecle qualquer tecla para continuar...')

	elif opcaoSelecionada == '7':
		print('Opção 7...')
		executa_procedure_sql_combinada('SP_PC_TEND_IGUAL_REAL_FIBRA_EMPRESARIAL',param)
		executa_procedure_sql_combinada('SP_PC_TEND_IGUAL_REAL_FIBRA_EMPRESARIAL_VL',param)
		
		executa_procedure_sql_combinada('SP_PC_TEND_IGUAL_REAL_FIBRA_VAREJO',param)
		executa_procedure_sql_combinada('SP_PC_TEND_IGUAL_REAL_FIBRA_VAREJO_VL',param)
		
		executa_procedure_sql_combinada('SP_PC_TEND_IGUAL_REAL_NOVA_FIBRA',param)
		executa_procedure_sql_combinada('SP_PC_TEND_IGUAL_REAL_NOVA_FIBRA_vl',param)

		executa_procedure_sql_combinada('SP_PC_TEND_IGUAL_REAL_TABELAS_FIBRA',param)
		executa_procedure_sql_combinada('SP_PC_TEND_IGUAL_REAL_TABELAS_FIBRA_VL',param)
		
		msg = f'Tendencias igualadas ao Realizado! : {datetime.today()}'
		enviaWhats(msg, num, key)
		
		a = input('Tecle qualquer tecla para continuar...')

	elif opcaoSelecionada == '8':
		print('Opção 8...')

		caminho_arquivo_sql = r'S:\\Resultados\\01_Relatorio Diario\\1 - Base Eventos\\02 - TENDÊNCIA\\ativar_tendencia_2023.sql'
		executa_arquivo_sql(caminho_arquivo_sql)
		
		executa_procedure_sql_combinada('SP_PC_BASES_SHAREPOINT',param)
		
		caminho_arquivo_sql_somente_fibra = r'S:\\Resultados\\01_Relatorio Diario\\1 - Base Eventos\\02 - TENDÊNCIA\\ativar_tendencia_somente_fibra_2023.sql'
		executa_arquivo_sql(caminho_arquivo_sql_somente_fibra)
				
		atualiza_TB_VALIDA_CARGA_TENDENCIA()
		msg = f'Bases liberadas para Sharepoint e PowerBi! : {datetime.today()}'
		enviaWhats(msg, num, key)
		
		a = input('Tecle qualquer tecla para continuar...')

	elif opcaoSelecionada == '9':
		print('Opção 9...')
		proc = 'SP_PC_Update_Ticket_Fibra_VAREJO_Tendencia_porRegiao'
		executa_procedure_sql_combinada(proc, param)
		proc = 'SP_PC_Update_Ticket_Fibra_EMPRESARIAL_Tendencia_porRegiao_IndCombo'
		executa_procedure_sql_combinada(proc, param)
		#proc = 'SP_PC_Update_Ticket_Fibra_VAREJO_DIARIO_porRegiao'
		#executa_procedure_sql_combinada(proc, param)
		#proc = 'SP_PC_Update_Ticket_Fibra_EMPRESARIAL_DIARIO_porRegiao_IndCombo'
		#executa_procedure_sql_combinada(proc, param)
		proc = 'SP_PC_TBL_RE_RELATORIO_RC_V2_TEND'
		executa_procedure_sql_combinada(proc, param)
		
		msg = f'Etapas de liberação do Ticket concluídas! : {datetime.today()}'
		enviaWhats(msg, num, key)

		a = input('Tecle qualquer tecla para continuar...')

	elif opcaoSelecionada == '10':
		print('Opção 10...')
		diretorio_arquivo_origem=r'S:\Resultados\01_Relatorio Diario\1 - Base Eventos\02 - TENDÊNCIA\TEND_FIBRA_e_UPDATEs.xlsx'
		diretorio_arquivo_destino=r'S:\Resultados\01_Relatorio Diario\1 - Base Eventos\02 - TENDÊNCIA\tend_email.xlsx'
		planilha_origem='UPDATE_TENDENCIA_VLL'
		copiar_colar_excel(diretorio_arquivo_origem, diretorio_arquivo_destino, planilha_origem)

		enviaEmailFimProcesso()
		a = input('Tecle qualquer tecla para continuar...')

	elif opcaoSelecionada == '11':
		print('Opção 11...Enviar PDVs OUTROS')
		EnviaPDVOutros()
		a = input('Tecle qualquer tecla para continuar...')

	elif opcaoSelecionada == '12':
		print('Opção 12...Gravando Historico de Tendencia CDO')
		executa_procedure_sql_combinada('dbo.SP_CDO_0003_TENDENCIA_PARA_HISTORICO',param)
		a = input('Tecle qualquer tecla para continuar...')

	elif opcaoSelecionada == '13':
		print('Opção 13...SAIR')
		break

	else:
		print('Opção Inválida')

print('FIM')