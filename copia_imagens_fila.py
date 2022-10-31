from playwright.sync_api import sync_playwright
import time


with sync_playwright() as p:
    navegador = p.chromium.launch(headless=False, slow_mo=100)
    pagina = navegador.new_page()
    #pagina.goto("https://tableau.oi.net.br/#/site/inteligenciadigital/views/NETADDS_16565284416200/NetAdds?:iid=1") 
    #pagina.goto("https://secure.oi.net.br/nidp/saml2/sso?id=OiPasswordClassCorporativoId&sid=0&option=credential&sid=0")
    pagina.goto("https://lobao.com/blog/") 

    print(pagina.title())
    


    #
    #pagina.locator('xpath = //*[@id="inputMatricula"]').fill('oi066724')
    #pagina.locator('xpath = //*[@id="inputPassword"]').fill('kMR2cB57')
    
    pagina.locator('xpath = /html/body/div/div/div/div/div/div[2]/div[1]/div').screenshot(path="screenshot.png")


    #pagina.wait_for_timeout(100000)
    navegador.close()
