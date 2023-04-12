import requests,segredos
from urllib.parse import quote
from datetime import date, datetime, timedelta


num= segredos.num
key = segredos.key


hoje = (datetime.today()- timedelta(days=0)).strftime('%d/%m/%Y') 
msg = f''''Olá Lobão ;)
Boa noite!
{hoje}
'''

#requests.get(
#    url=f'https://api.callmebot.com/whatsapp.php?phone={numero}&text={quote(mensagem)}&apikey={apikey}'
#)

def enviaWhats (mensagem, numero, apikey):
    requests.get(
    url=f'https://api.callmebot.com/whatsapp.php?phone={numero}&text={quote(mensagem)}&apikey={apikey}'
    )

enviaWhats(msg, num, key)