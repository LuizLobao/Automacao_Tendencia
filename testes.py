import segredos
import pandas as pd
import pyodbc
import time
from datetime import date, datetime
import win32com.client as win32

hoje = datetime.today().strftime('%d/%m/%Y')
AAAAMMDD = datetime.today().strftime('%Y%m%d')
print(hoje)

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

enviaEmaileAnexo()

 