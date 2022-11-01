from html2image import Html2Image
hti = Html2Image(browser = 'chromium')

#hti.screenshot(url='https://www.lobao.com', save_as='python_org.png')

pagina = 'https://www.lobao.com'
html = '//*[@id="post-9"]/div/div[2]/div/div[2]/div/div[8]/div/p'
hti.screenshot(url=pagina, html_str=html, save_as='teste.png')   