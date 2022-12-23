from playwright.sync_api import Page, sync_playwright, expect
import time
from datetime import date, datetime, timedelta
import os, zipfile
import subprocess
from tqdm import tqdm


subprocess.run('cls', shell=True)


hoje = date.today().strftime('%d/%m/%Y')
ontem = (date.today() - timedelta(days=1)).strftime('%Y%m%d')
AAAAMMDD = datetime.today().strftime('%Y%m%d')
AnoMes = datetime.today().strftime('%Y%m')


print(f'Hoje: {hoje}')
print(f'Ontem: {ontem}')

def download_bov(pagina_bov):
	dicio={}
	with sync_playwright() as p:

		navegador = p.chromium.launch(headless=True)
		pagina = navegador.new_page()
		
		with tqdm(total=35) as barra_progresso:
			for relatorio, links in pagina_bov.items():
				pagina.goto(links)
				linha = 1
				#print(f'Verificando relatório: {relatorio}')
				while linha <= 5:
						#print(f'Verificando linha: {linha}')
						id=(pagina.locator(f'xpath = //*[@id="lbl1"]/table/tbody/tr[{linha}]/td[1]').text_content()).replace(u'\xa0',u'')
						filtro=(pagina.locator(f'xpath = //*[@id="lbl1"]/table/tbody/tr[{linha}]/td[3]').text_content()).replace(u'\xa0',u'')
						de=(pagina.locator(f'xpath = //*[@id="lbl1"]/table/tbody/tr[{linha}]/td[5]').text_content()).replace(u'\xa0',u'')
						ate=(pagina.locator(f'xpath = //*[@id="lbl1"]/table/tbody/tr[{linha}]/td[6]').text_content()).replace(u'\xa0',u'')
						usuario=(pagina.locator(f'xpath = //*[@id="lbl1"]/table/tbody/tr[{linha}]/td[7]').text_content()).replace(u'\xa0',u'')
						criacao=(pagina.locator(f'xpath = //*[@id="lbl1"]/table/tbody/tr[{linha}]/td[8]').text_content()).replace(u'\xa0',u'')
						datacriacao =criacao.split(" ")[0]
						if usuario == 'AUTO' and filtro == 'N' and ate == ontem :
							dicio.update({relatorio:(links, linha)})
							with pagina.expect_download() as download_info:
								pagina.locator(f'//*[@id="lbl1"]/table/tbody/tr[{linha}]/td[10]/font/b/a/img').click()
								download = download_info.value
								#print(download.path())
								nomearq = download.suggested_filename
								download.save_as(fr'C:\\Users\\oi066724\\Downloads\\BOVs\\{nomearq}')
						linha += 1
						barra_progresso.update(1)
				#barra_progresso.update(1)
			#print(dicio)
		navegador.close()
	
def unzip(dir_origem, dir_destino):
	print('INICIANDO UNZIP)')
	dir_zip = dir_origem
	dir_unzip = dir_destino
	extension = '.ZIP'
	
	os.chdir(dir_zip)
	with tqdm(total=7) as barra_progresso:
		for item in os.listdir(dir_zip):
			if item.endswith(extension):
				file_name = os.path.abspath(item)
				zip_ref = zipfile.ZipFile(file_name)
				zip_ref.extractall(dir_unzip)
				zip_ref.close()
				os.remove(file_name)
			barra_progresso.update(1)	



download_bov({'1067':'https://portalbi.telemar/AdminRelBatchCadastroJETL.aspx?idRelatorio=1067&idAplicacao=1'
,'1058':'https://portalbi.telemar/AdminRelBatchCadastroJETL.aspx?idRelatorio=1058&idAplicacao=1'
,'1059':'https://portalbi.telemar/AdminRelBatchCadastroJETL.aspx?idRelatorio=1059&idAplicacao=1'
,'1065':'https://portalbi.telemar/AdminRelBatchCadastroJETL.aspx?idRelatorio=1065&idAplicacao=1'
,'1064':'https://portalbi.telemar/AdminRelBatchCadastroJETL.aspx?idRelatorio=1064&idAplicacao=1'
,'6162':'https://portalbi.telemar/AdminRelBatchCadastroJETL.aspx?idRelatorio=6162&idAplicacao=11'
,'6163':'https://portalbi.telemar/AdminRelBatchCadastroJETL.aspx?idRelatorio=6163&idAplicacao=11'
})

unzip('C:\\Users\\oi066724\\Downloads\\BOVs\\','C:\\Users\\oi066724\\Downloads\\BOVs\\unzip\\')