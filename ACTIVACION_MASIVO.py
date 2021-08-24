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

#for i in range(1,18):
fecha="ACTIVACION_MASIVO_ENE"+str(time.strftime("%d%m%y"))+".xlsx"
#fecha="ACTIVACION_MASIVO_ENE"+str(i)+"012021.xlsx"

    ##EXTRACCION DE LA DATA
    ####################
driver=webdriver.Chrome(executable_path=r'./chromedriver.exe')

driver.get('https://recargasprepago.claro.cr:4448/pretups/')

txt_usuario='#loginID'
txt_contraseña='#password'
btn_entrar_sistema='body > form:nth-child(2) > table:nth-child(6) > tbody > tr > td > table:nth-child(2) > tbody > tr:nth-child(5) > td.tabcol > input:nth-child(1)'

driver.find_element_by_css_selector(txt_usuario).send_keys('DISPIKICR')
driver.find_element_by_css_selector(txt_contraseña).send_keys('PIKI79/')
driver.find_element_by_css_selector(btn_entrar_sistema).click()

time.sleep(3)

    ############################################################
a= driver.find_elements_by_tag_name('frame')
driver.switch_to.frame(a[0])
    ############################################################

driver.find_element_by_link_text('Channel reports-C2S').click()
time.sleep(5)
    ### seleccionar jerarquia

element = driver.find_element_by_name("serviceType")
time.sleep(3)

webElem=driver.find_element_by_xpath("//*[(@value = 'ALL')]")
webElem.click()
time.sleep(3)

select=driver.find_element_by_name("transferStatus")
select.click()
time.sleep(3)

webElem=driver.find_element_by_xpath("//*[(@value = '200')]")
webElem.click()
time.sleep(3)

fechainicio=driver.find_element_by_name("currentDate")
fecha_a=datetime.strftime((datetime.today() - timedelta(days=1)),'%d/%m/%y')
    #datetime.strftime((datetime.today() - timedelta(days=1)),'%d/%m/%y')
    #if i<10:
    #    f="0"+str(i)+"/01/21"
    #else:
    #    f=""+str(i)+"/01/21"
    
fechainicio.send_keys(fecha_a)
time.sleep(5)

hora_inicial="00:00"
hora_final="23:59"

fechafin=driver.find_element_by_name("fromTime")
fechafin.send_keys(hora_inicial)
time.sleep(3)


fechafin=driver.find_element_by_name("toTime")
fechafin.send_keys(hora_final)
time.sleep(3)


btnE=driver.find_element_by_name("submitButton")
btnE.click()
time.sleep(3)
    #get current window handle
p = driver.current_window_handle

    #get first child window#get first child window
chwd = driver.window_handles
driver.switch_to.window(chwd[1])
driver.maximize_window()
print("Child window title: " + driver.title)

time.sleep(3)

saverpt=driver.find_element_by_id("save")
saverpt.click()
time.sleep(3)

saverpt =driver.find_element_by_xpath("//*[@id='__menuBar']/div[6]/ul/li[1]/div")
saverpt.click()
time.sleep(3)

saverpt =driver.find_element_by_xpath("//*[@id='__menuBar']/div[6]/ul/li[1]/div/ul/li[4]")
saverpt.click()
time.sleep(3)

saverpt =driver.find_element_by_id("ok")
saverpt.click()

time.sleep(10)



    ##MODELADO DE LA DATA
    ##########################################################################
    # #\\192.168.0.85\Users\50576\Desktop\python
p_open=pd.read_excel(r'C:\Users\50576\Downloads\c2sTransferChannelUserNew.xlsx')

p_open_=pd.DataFrame(p_open)
p_open_=p_open_.drop([0,1,2,3,4,5,6,7,8,9,10,11,12])
p_open_=p_open_.drop(['Unnamed: 0', 'Unnamed: 1', 'Unnamed: 3','Unnamed: 9','Unnamed: 17','Unnamed: 18'], axis=1)
p_open_= p_open_.dropna(axis=0, subset=['Unnamed: 2'])
p_open_.rename(columns={'Unnamed: 2':"transferencia",
'Unnamed: 4':"tipo",
'Unnamed: 5':"cliente",
'Unnamed: 6':"pos",
'Unnamed: 7':"service",
'Unnamed: 8':"pos_final",
'Unnamed: 13':"monto"
},inplace=True)
p_open_.reset_index(inplace=True,drop=True) 

p_open_=pd.DataFrame(p_open_[
    ["transferencia",
    "tipo",
    "cliente",
    "pos",  
    "service",
    "pos_final",
    "monto"]])

p_open_['fecha'] = p_open_['transferencia'].str.slice(1, 7)

p_open_.to_excel(r'C:\Users\50576\Documents\ARCHIVOS BI\TAE CR\ACTIVACION\%s' %fecha,engine='xlsxwriter',
index=False,sheet_name='Pagos Open')

shutil.move(r'C:\Users\50576\Downloads\c2sTransferChannelUserNew.xlsx', r'C:\Users\50576\AppData\Local\Temp\data.xlsx')

time.sleep(3)
driver.quit()


    #print(p_open_)