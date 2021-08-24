#import numpy as np
import smtplib,ssl 
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText 
from email.mime.base import MIMEBase 
from email import encoders
from sqlalchemy import create_engine
from selenium import webdriver

from datetime import datetime, timedelta
from selenium.webdriver.remote.switch_to import SwitchTo

import shutil
import os
import glob
import time

import email
import re   
from click import style
import html
from smtplib import SMTPAuthenticationError
from email.mime.image import MIMEImage

from selenium.webdriver.common import by
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import pandas as pd

#establece parametros al cargar el driver de chrome
driver=webdriver.Chrome(executable_path=r'./chromedriver.exe')


fecha_inicial="01/08/20"
hora_inicial="00:00"
hora_final="23:59"
fecha_aterior=datetime.strftime((datetime.today() - timedelta(days=1)),'%d/%m/%y')

#establece credenciales de pretups
email="dist_inverfinanc"
senha="Claro#im2022"


driver.get('https://recargas.claro.com.ni:4446/pretups/')

txt_usuario='//*[@id="loginID"]'
txt_contraseña='//*[@id="password"]'
btn_entrar_sistema='/html/body/form[2]/table[3]/tbody/tr/td/table[2]/tbody/tr[5]/td[2]/input[1]'
time.sleep(4)
driver.find_element_by_xpath(txt_usuario).send_keys(email)
time.sleep(3)
driver.find_element_by_xpath(txt_contraseña).send_keys(senha)
time.sleep(3)
driver.find_element_by_xpath(btn_entrar_sistema).click()
time.sleep(3)

##########################################################
a=driver.find_elements_by_tag_name('frame')
driver.switch_to.frame(a[0])
time.sleep(3)
##########################################################
driver.find_element_by_link_text('Channel reports-C2S').click()
time.sleep(3)
###############################

# seleccionar jerarquia
select=driver.find_element_by_name("serviceType")
select.click()
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
fechainicio.send_keys(fecha_aterior)
time.sleep(3)

fechafin=driver.find_element_by_name("fromTime")
fechafin.send_keys(hora_inicial)
time.sleep(10)

fechafin=driver.find_element_by_name("toTime")
fechafin.send_keys(hora_final)
time.sleep(10)

btngenrpt=driver.find_element_by_name("submitButton")
btngenrpt.click()
time.sleep(30)

p = driver.current_window_handle
#get first child window
chwd = driver.window_handles
driver.switch_to.window(chwd[1])
driver.maximize_window()
print("Child window title: " + driver.title)

time.sleep(10)

saverpt=driver.find_element_by_id("save")
saverpt.click()
time.sleep(10)

saverpt =driver.find_element_by_xpath("//*[@id='__menuBar']/div[6]/ul/li[1]/div")
saverpt.click()
time.sleep(10)

saverpt =driver.find_element_by_xpath("//*[@id='__menuBar']/div[6]/ul/li[1]/div/ul/li[4]")
saverpt.click()
time.sleep(10)

saverpt =driver.find_element_by_id("ok")
saverpt.click()

time.sleep(15)

driver.quit()

time.sleep(5)

DATABASES = {
    'apoyo':{
        'NAME': 'apoyo',
        'USER': 'kenth',
        'PASSWORD': 'K%8n)th#2020*',
        'HOST': 'db.multipagos.net',
        'PORT': 5432,
    },
}


db = DATABASES['apoyo']

#construct an engine connection string
engine_string = "postgresql+psycopg2://{user}:{password}@{host}:{port}/{database}".format(
    user = db['USER'],
    password = db['PASSWORD'],
    host = db['HOST'],
    port = db['PORT'],
    database = db['NAME'],
)

engine = create_engine(engine_string)

##############################################################

p_open=pd.read_excel(r'C:\Users\50576\Downloads\c2sTransferChannelUserNew.xlsx')
p_open2=pd.read_excel(r'C:\Users\50576\Documents\Multipagos\Pago_Pretups\COBRADORES.xlsx')

p_open_=pd.DataFrame(p_open)
p_open_=p_open_.drop([0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16])
p_open_=p_open_.drop(['Unnamed: 0', 'Unnamed: 1', 'Unnamed: 3','Unnamed: 9','Unnamed: 17','Unnamed: 18'], axis=1)
p_open_= p_open_.dropna(axis=0, subset=['Unnamed: 2'])
p_open_.rename(columns={'Unnamed: 2':"codigo",
'Unnamed: 4':"requerimiento",
'Unnamed: 5':"usuario",
'Unnamed: 6':"pos",
'Unnamed: 7':"emisor",
'Unnamed: 8':"destinatario",
'Unnamed: 10':"tipo_servicio",
'Unnamed: 11':"servicio",
'Unnamed: 12':"sub_servicio",
'Unnamed: 13':"monto",
'Unnamed: 14':"credito",
'Unnamed: 15':"bono",
'Unnamed: 16':"procesamiento"
},inplace=True)

p_open2= pd.DataFrame(p_open2, columns = ['pos', 'sucursal'])
p_open2["pos"]=p_open2["pos"].astype(str)

df_merge=pd.merge(p_open_, p_open2, on="pos", how="left")
now=datetime.now()
date_time = now.strftime("%Y-%m-%d %H:%M:%S")
df_merge['fecha_registro']= date_time

df_merge.to_sql('apoyo_web_prectus', engine,if_exists='append',index=False)



shutil.move(r'C:\Users\50576\Downloads\c2sTransferChannelUserNew.xlsx', r'C:\Users\50576\AppData\Local\Temp\activacion.xlsx')

print(p_open_)