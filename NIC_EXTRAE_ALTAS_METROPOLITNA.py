from  selenium import webdriver
import time
from selenium.webdriver.common import by
from datetime import datetime, timedelta
from selenium.webdriver.remote.switch_to import SwitchTo
from selenium.webdriver.common.by import By
import pandas as pd
from datetime import datetime, timedelta
import shutil
import os
import glob
import time

driver=webdriver.Chrome(executable_path=r'./chromedriver.exe')

driver.get('http://172.26.59.95/directorio_indicadores/directorio_indicadores/directorio.aspx#no-back-button')
time.sleep(10)
driver.find_element_by_xpath('//*[@id="ctl01"]/div[4]/a[1]/div').click()
time.sleep(10)
driver.find_element_by_xpath('//*[@id="TxtUsuario"]').send_keys('METRO001')
time.sleep(10)
driver.find_element_by_xpath('//*[@id="TxtPass"]').send_keys('123')
time.sleep(10)
driver.find_element_by_xpath('//*[@id="btnlogin"]').click()

time.sleep(19)

driver.find_element_by_xpath('//*[@id="LeftPanel_nbMain_GHC0"]').click()
time.sleep(10)
driver.find_element_by_xpath('//*[@id="LeftPanel_nbMain_I0i0_"]').click()
time.sleep(10)
driver.find_element_by_xpath('//*[@id="EmployeeSelectorPanel_MainContent_btnExportar"]').click()

############################################################

time.sleep(12)

driver.quit()

fecha="METROPOLITANA_Detalle"+str(time.strftime("%m%y"))+".csv"

p_open=pd.read_csv('C:/Users/50576/Downloads/METROPOLITANA_Detalle.csv', sep='\t',encoding = 'utf-8',error_bad_lines=False)
p_open.to_csv(r'C:\Users\50576\Documents\PREPAGO\%s'%fecha,sep='\t',encoding = 'utf-8',
index=False)

shutil.move(r'C:\Users\50576\Downloads\METROPOLITANA_Detalle.csv', r'C:\Users\50576\Documents\COPIA ALTA MES CORRIENTE PREPAGO\alta.csv')