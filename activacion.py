from  selenium import webdriver
import time
from selenium.webdriver.common import by
from datetime import datetime, timedelta
from selenium.webdriver.remote.switch_to import SwitchTo
from selenium.webdriver.common.by import By

driver=webdriver.Chrome(executable_path=r'./chromedriver.exe')
driver.get('https://recargasprepago.claro.cr:4448/pretups/')

txt_usuario='#loginID'
txt_contraseña='#password'
btn_entrar_sistema='body > form:nth-child(2) > table:nth-child(6) > tbody > tr > td > table:nth-child(2) > tbody > tr:nth-child(5) > td.tabcol > input:nth-child(1)'

driver.find_element_by_css_selector(txt_usuario).send_keys('MULPIKICR')
driver.find_element_by_css_selector(txt_contraseña).send_keys('PIKI27/')
driver.find_element_by_css_selector(btn_entrar_sistema).click()

time.sleep(3)

############################################################
a= driver.find_elements_by_tag_name('frame')
driver.switch_to.frame(a[0])
############################################################

driver.find_element_by_link_text('Channel reports-C2S').click()
time.sleep(2)
### seleccionar jerarquia

element = driver.find_element_by_name("serviceType")
time.sleep(1)

webElem=driver.find_element_by_xpath("//*[(@value = 'ALL')]")
webElem.click()
time.sleep(1)

select=driver.find_element_by_name("transferStatus")
select.click()
time.sleep(3)

webElem=driver.find_element_by_xpath("//*[(@value = '200')]")
webElem.click()
time.sleep(3)

fechainicio=driver.find_element_by_name("currentDate")
fecha_a=datetime.strftime((datetime.today() - timedelta(days=1)),'%d/%m/%y')
fechainicio.send_keys(fecha_a)

time.sleep(3)
hora_inicial="00:00"
hora_final="23:59"
Hi=driver.find_element_by_name("fromTime")
Hi.send_keys(hora_inicial)
time.sleep(2)

Hf=driver.find_element_by_name("toTime")
Hf.send_keys(hora_final)
time.sleep(2)

btngenrpt=driver.find_element_by_name("submitButton")
btngenrpt.click()
time.sleep(2)

#get current window handle
p = driver.current_window_handle

#get first child window#get first child window
chwd = driver.window_handles
driver.switch_to.window(chwd[1])
driver.maximize_window()
print("Child window title: " + driver.title)

time.sleep(2)

saverpt=driver.find_element_by_id("save")
saverpt.click()
time.sleep(2)

saverpt =driver.find_element_by_xpath("//*[@id='__menuBar']/div[6]/ul/li[1]/div")
saverpt.click()
time.sleep(2)

saverpt =driver.find_element_by_xpath("//*[@id='__menuBar']/div[6]/ul/li[1]/div/ul/li[4]")
saverpt.click()
time.sleep(2)

saverpt =driver.find_element_by_id("ok")
saverpt.click()

time.sleep(10)

driver.quit()


