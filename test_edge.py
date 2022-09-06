from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep

driver = webdriver.Edge(r'C:\\Users\\oi066724\\Documents\\Python\\Python\\msedgedriver.exe')
driver.get('http://10.20.83.116/aplicacao/monitor/')

    # /html/body/table[2]/tbody/tr[7]/td[8] #BOV_1066 hora fim carga
    # /html/body/table[2]/tbody/tr[8]/td[8] #BOV_1067 hora fim carga
    # /html/body/table[2]/tbody/tr[9]/td[8] #BOV_1059 hora fim carga
    # /html/body/table[2]/tbody/tr[10]/td[8] #BOV_1064 hora fim carga
    # /html/body/table[2]/tbody/tr[11]/td[8] #BOV_1065 hora fim carga
    # /html/body/table[2]/tbody/tr[12]/td[8] #BOV_1058 hora fim carga

    # /html/body/table[2]/tbody/tr[82]/td[8] #6162 hora fim carga
    # /html/body/table[2]/tbody/tr[83]/td[8] #6163 hora fim carga

sleep(3)

val_1066 = driver.find_element(By.XPATH,"/html/body/table[2]/tbody/tr[7]/td[8]").text
val_1067 = driver.find_element(By.XPATH,"/html/body/table[2]/tbody/tr[8]/td[8]").text
val_1059 = driver.find_element(By.XPATH,"/html/body/table[2]/tbody/tr[9]/td[8]").text
val_1064 = driver.find_element(By.XPATH,"/html/body/table[2]/tbody/tr[10]/td[8]").text
val_1065 = driver.find_element(By.XPATH,"/html/body/table[2]/tbody/tr[11]/td[8]").text
val_1058 = driver.find_element(By.XPATH,"/html/body/table[2]/tbody/tr[12]/td[8]").text

val_6162 = driver.find_element(By.XPATH, "/html/body/table[2]/tbody/tr[82]/td[8]").text
val_6163 = driver.find_element(By.XPATH, "/html/body/table[2]/tbody/tr[82]/td[8]").text

print(f'Rel 1066: {val_1066}')
print(f'Rel 1067: {val_1067}')
print(f'Rel 1059: {val_1059}')
print(f'Rel 1064: {val_1064}')
print(f'Rel 1065: {val_1065}')
print(f'Rel 1058: {val_1058}')
print(f'Rel 6162: {val_6162}')
print(f'Rel 6163: {val_6163}')

