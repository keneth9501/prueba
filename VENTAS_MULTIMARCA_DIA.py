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
#import datetime
#mydate = datetime.datetime.now() }
fecha_=datetime.strftime((datetime.today() + timedelta(days=1)),'%d/%m/%y')
#g=mydate.strftime("%B").upper()

fecha="venta_MULTIMARCA_ENE"+str(time.strftime("%d%m%y"))+".xlsx"
##EXTRACCION DE LA DATA

####################
driver=webdriver.Chrome(executable_path=r'./chromedriver.exe')
driver.get('https://recargasprepago.claro.cr:4448/pretups/')
txt_usuario='#loginID'
txt_contraseña='#password'
btn_entrar_sistema='body > form:nth-child(2) > table:nth-child(6) > tbody > tr > td > table:nth-child(2) > tbody > tr:nth-child(5) > td.tabcol > input:nth-child(1)'

driver.find_element_by_css_selector(txt_usuario).send_keys('MULPIKICR')
driver.find_element_by_css_selector(txt_contraseña).send_keys('PIKI27/')
driver.find_element_by_css_selector(btn_entrar_sistema).click()

time.sleep(16)

############################################################
a= driver.find_elements_by_tag_name('frame')
driver.switch_to.frame(a[0])
############################################################

driver.find_element_by_link_text('Channel reports-C2C').click()
time.sleep(10)
### seleccionar jerarquia

element = driver.find_element_by_name("txnSubType")
time.sleep(10)

webElem=driver.find_element_by_xpath("//*[(@value = 'T')]")
webElem.click()
time.sleep(10)

select=driver.find_element_by_name("transferInOrOut")
select.click()
time.sleep(10)

webElem=driver.find_element_by_xpath("//*[(@value = 'IN')]")
webElem.click()
time.sleep(10)

fechainicio=driver.find_element_by_name("fromDate")
fecha_a=datetime.strftime((datetime.today() - timedelta(days=1)),'%d/%m/%y')
#datetime.strftime((datetime.today() - timedelta(days=1)),'%d/%m/%y')
fechainicio.send_keys(fecha_a)

time.sleep(10)
fecha_i=driver.find_element_by_name("toDate")
fecha_at =datetime.strftime((datetime.today() - timedelta(days=1)),'%d/%m/%y')
#datetime.strftime((datetime.today() - timedelta(days=1)),'%d/%m/%y')
fecha_i.send_keys(fecha_at)

btnE=driver.find_element_by_name("c2cTrfRetWid")
btnE.click()

time.sleep(20)
#get current window handle
p = driver.current_window_handle

#get first child window#get first child window
chwd = driver.window_handles
driver.switch_to.window(chwd[1])
driver.maximize_window()

print("Child window title: " + driver.title)
time.sleep(30)

saverpt=driver.find_element_by_id("save")
saverpt.click()
time.sleep(9)

saverpt =driver.find_element_by_xpath("//*[@id='__menuBar']/div[6]/ul/li[1]/div")
saverpt.click()
time.sleep(9)

saverpt =driver.find_element_by_xpath("//*[@id='__menuBar']/div[6]/ul/li[1]/div/ul/li[4]")
saverpt.click()
time.sleep(9)

saverpt =driver.find_element_by_id("ok")
saverpt.click()

time.sleep(9)





##MODELADO DE LA DATA
##########################################################################
# #\\192.168.0.85\Users\50576\Desktop\python
p_open=pd.read_excel(r'C:\Users\50576\Downloads\C2cRetWidTransferChannelUser.xlsx')
p_open_=pd.DataFrame(p_open)
p_open_=p_open_.drop([0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16])
p_open_=p_open_.drop(['Unnamed: 0', 'Unnamed: 2', 'Unnamed: 6','Unnamed: 8','Unnamed: 10'], axis=1)
p_open_= p_open_.dropna(axis=0, subset=['Unnamed: 1'])
p_open_.rename(columns={'Unnamed: 4':"origen",
'Unnamed: 19':"monto_transferencia",
'Unnamed: 15':"fecha",
'Unnamed: 11':"transferencia",
'Unnamed: 9':"pos_destino",
'Unnamed: 5':"pos_origen",
'Unnamed: 7':"destino"
},inplace=True)
p_open_.reset_index(inplace=True,drop=True) 

p_open_=pd.DataFrame(p_open_[
    ["origen",
    "pos_origen",
    "destino",
    "pos_destino",  
    "transferencia",
    "fecha",
     "monto_transferencia"]])

p_open_.to_excel(r'C:\Users\50576\Documents\ARCHIVOS BI\TAE CR\COLOCACION\%s'%fecha,engine='xlsxwriter',
index=False,sheet_name='pagos')

shutil.move(r'C:\Users\50576\Downloads\C2cRetWidTransferChannelUser.xlsx', r'C:\Users\50576\AppData\Local\Temp\data.xlsx')
time.sleep(3)
driver.quit()