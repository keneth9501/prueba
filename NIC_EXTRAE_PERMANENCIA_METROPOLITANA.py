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

time.sleep(8)
driver.find_element_by_xpath('//*[@id="ctl01"]/div[4]/a[1]/div').click()
time.sleep(8)
driver.find_element_by_xpath('//*[@id="TxtUsuario"]').send_keys('METRO001')
time.sleep(8)
driver.find_element_by_xpath('//*[@id="TxtPass"]').send_keys('123')
time.sleep(8)
driver.find_element_by_xpath('//*[@id="btnlogin"]').click()
time.sleep(10)

driver.find_element_by_xpath('//*[@id="LeftPanel_nbMain_GHC1"]').click()
time.sleep(8)

driver.find_element_by_xpath('//*[@id="LeftPanel_nbMain_I1i0_T"]').click()
time.sleep(8)
driver.find_element_by_xpath('//*[@id="EmployeeSelectorPanel_MainContent_btnexportar"]').click()

###########################################################

time.sleep(140)

driver.quit()

#MODELADO DE LA DATA
#########################################################################

df=pd.read_csv('C:/Users/50576/Downloads/METROPOLITANA_Detalle.csv', sep='\t',encoding = 'utf-8',error_bad_lines=False)
time.sleep(8)
p_open=pd.read_excel(r'C:\Users\50576\Documents\PERMANENCIA\PERMANENCIA.xlsx')
time.sleep(60)
a=df['IDPERIODO'].max()
df.insert(0,"PERIODO",a,True)

aa=p_open['PERIODO'].max()

if(aa==a):
    p_open.drop(p_open[p_open['PERIODO']==aa].index, inplace = True)
    p_open=pd.concat([p_open,df])
elif(df.empty):
    print('No hay registro')
else:
    p_open=pd.concat([p_open,df])  

p_open.reset_index(inplace=True,drop=True) 

fecha='PERMANENCIA.xlsx'

shutil.move(r'C:\Users\50576\Documents\PERMANENCIA\PERMANENCIA.xlsx', r'C:\Users\50576\Documents\PERMANENCIA\Permanencia TEMP\PERMANENCIA.xlsx')

p_open.to_excel(r'C:\Users\50576\Documents\PERMANENCIA\%s' %fecha,engine='xlsxwriter',
index=False,sheet_name='PERMANENCIA')
time.sleep(60)
shutil.move(r'C:\Users\50576\Downloads\METROPOLITANA_Detalle.csv', r'C:\Users\50576\AppData\Local\Temp\data.csv')