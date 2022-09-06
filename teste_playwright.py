#antes de começar:
#pip install playwright
# playwright install
# https://www.youtube.com/watch?v=1NNMzL4W8ws&t=137s


#usar a ferramenta PLAYWRIGHT para navegar na WEB (sem janela) para puxar as infos necessárias
from playwright.sync_api import sync_playwright
import time

with sync_playwright() as p:
    navegador = p.chromium.launch(headless=True)
    pagina = navegador.new_page()
    pagina.goto("https://www.google.com/search?q=cota%C3%A7%C3%A3o+dolar")
    print(pagina.title())

    dolar = pagina.locator('xpath = //*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').text_content()
    print(f'Dolar = {dolar}')
    
    pagina.goto("https://www.google.com/search?q=cota%C3%A7%C3%A3o+euro")
    euro = pagina.locator('xpath = //*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').text_content()
    print(f'Euro = {euro}')

    #time.sleep(15)
    navegador.close()


#NO MONITOR DE CARGAS 
#BOV 1066
#BOV 1067
#BOV 1059
#BOV 1065
#BOV 1064
#BOV 1058
#BOV 6162
#BOV 6163

#NO SHAREPOINT DE TV ANTENEIROS

