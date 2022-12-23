import segredos
from playwright.sync_api import sync_playwright
from datetime import date, datetime
from tqdm import tqdm
import time



hoje = datetime.today().strftime('%d/%m/%Y')
AAAAMMDD = datetime.today().strftime('%Y%m%d')
print(hoje)



def puxa_dts_cargas():
	dicio = {"arquivo":"data","HOJE":hoje}
	with sync_playwright() as p:

		navegador = p.chromium.launch(headless=True)
		pagina = navegador.new_page()
		pagina.goto("http://10.20.83.116/aplicacao/monitor/")
		linha = 1
		with tqdm(total=95) as barra_progresso:
			while linha <= 95:
					arquivo=(pagina.locator(f'xpath = //*[@id="mytable"]/tbody/tr[{linha}]/td[9]').text_content())
					DataFim=(pagina.locator(f'xpath = //*[@id="mytable"]/tbody/tr[{linha}]/td[8]').text_content())
					Status=(pagina.locator(f'xpath = //*[@id="mytable"]/tbody/tr[{linha}]/td[6]').text_content())
					#print(arquivo)
					if Status == 'Carga realizada':
						dicio.update({arquivo:DataFim})
					linha += 1
					barra_progresso.update(1)
		navegador.close()
	
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



puxa_dts_cargas()




