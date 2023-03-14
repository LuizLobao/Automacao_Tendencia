import os
import win32com.client as win32
from PIL import ImageGrab
import segredos
from datetime import date, datetime, timedelta

hoje = (datetime.today()- timedelta(days=0)).strftime('%d/%m/%Y') 

def enviaEmailFimProcesso():		

	workbook_path = r'S:\Resultados\01_Relatorio Diario\1 - Base Eventos\02 - TENDÊNCIA\TEND_FIBRA_e_UPDATEs.xlsx'
	excel = win32.Dispatch('Excel.Application')

	wb = excel.Workbooks.Open(workbook_path)
	sheet = wb.Sheets.Item(3)
	excel.visible = 1

	copyrange = sheet.Range('A23:D48')

	copyrange.CopyPicture(Appearance=1, Format=2)
	ImageGrab.grabclipboard().save('paste.png')
	excel.Quit()

	image_path = r'C:\Users\oi066724\Documents\Python\Automacao_Tendencia\paste.png'

	# create a html body template and set the **src** property with `{}` so that we can use
	# python string formatting to insert a variable
	html_body = """
		<div>
			Tendências Liberadas no banco de dados INFOCANA.
		</div>
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

	# display the message to review
	message.Display()

	# save or send the message
	message.Send()
	print("Email Enviado")

enviaEmailFimProcesso()    