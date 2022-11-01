from playwright.sync_api import sync_playwright
import time
import segredos


with sync_playwright() as p:
    navegador = p.chromium.launch(headless=False)#kMR2cB57, slow_mo=100)
    pagina = navegador.new_page()
    #pagina.goto("https://tableau.oi.net.br/#/site/inteligenciadigital/views/NETADDS_16565284416200/NetAdds?:iid=1") 
    pagina.goto("https://tableau.oi.net.br/#/site/inteligenciadigital/views/PainisOperacionaisdeVendas/Diarioesemanal?:iid=1")
    #pagina.goto("https://lobao.com/blog/") 


    


    #
    pagina.locator('xpath = //*[@id="inputMatricula"]').fill(segredos.meu_id)
    pagina.locator('xpath = //*[@id="inputPassword"]').fill(segredos.senha_rede)
    
    time.sleep(30)

    #pagina.goto("https://tableau.oi.net.br/#/site/inteligenciadigital/views/PainisOperacionaisdeVendas/Diarioesemanal?:iid=1")
    print(pagina.title())
    pagina.locator('#tabZoneId951 > div:nth-child(1)').screenshot(path="screenshot2.png")
    

    #pagina.locator('xpath = //*[@id="tabZoneId975"]/div"]').focus()
    #navegador.close()


