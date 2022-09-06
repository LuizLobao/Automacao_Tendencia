from playwright.sync_api import sync_playwright

def puxa_dts_cargas():
    with sync_playwright() as p:
        navegador = p.chromium.launch(headless=True)
        pagina = navegador.new_page()
        pagina.goto("http://10.20.83.116/aplicacao/monitor/")
        
        dt1066 = pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[7]/td[8]').text_content()
        dt1067 = pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[8]/td[8]').text_content()
        dt1064 = pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[9]/td[8]').text_content()
        dt1059 = pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[10]/td[8]').text_content()
        dt1065 = pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[11]/td[8]').text_content()
        dt1058 = pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[12]/td[8]').text_content()
        dt6163 = pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[82]/td[8]').text_content()
        dt6162 = pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[83]/td[8]').text_content()

        dw1066 = pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[7]/td[5]').text_content()
        dw1067 = pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[8]/td[5]').text_content()
        dw1064 = pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[9]/td[5]').text_content()
        dw1059 = pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[10]/td[5]').text_content()
        dw1065 = pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[11]/td[5]').text_content()
        dw1058 = pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[12]/td[5]').text_content()
        dw6163 = pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[82]/td[5]').text_content()
        dw6162 = pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[83]/td[5]').text_content()

        di1066 = pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[7]/td[7]').text_content()
        di1067 = pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[8]/td[7]').text_content()
        di1064 = pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[9]/td[7]').text_content()
        di1059 = pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[10]/td[7]').text_content()
        di1065 = pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[11]/td[7]').text_content()
        di1058 = pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[12]/td[7]').text_content()
        di6163 = pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[82]/td[7]').text_content()
        di6162 = pagina.locator('xpath = //*[@id="mytable"]/tbody/tr[83]/td[7]').text_content()

        navegador.close()

        datas_fim = [dt1066,dt1067,dt1064,dt1059,dt1065,dt1058,dt6163,dt6162]
        datas_down = [dw1066,dw1067,dw1064,dw1059,dw1065,dw1058,dw6163,dw6162]
        datas_ini = [di1066,di1067,di1064,di1059,di1065,di1058,di6163,di6162]

        return datas_fim, datas_down, datas_ini