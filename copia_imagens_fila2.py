import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import os


driver = webdriver.Edge()
driver.get("https://tableau.oi.net.br/#/site/inteligenciadigital/views/NETADDS_16565284416200/NetAdds?:iid=1")
time.sleep(30)


driver.get("https://tableau.oi.net.br/#/site/inteligenciadigital/views/NETADDS_16565284416200/NetAdds?:iid=2")
time.sleep(10)
body = driver.find_element(By.XPATH, '/html/body/div[2]/div[2]/div[2]/div[1]/div/div[2]/div[44]/div')

body.screenshot('retirada.png')


driver.close()