import requests
from requests.auth import HTTPBasicAuth
usuario="keneth.morales"
contrasena="Alacran2020"

#url1 = "http://sys.multipagos.net/apoyo/API/pagos/"
url2 = "http://sys.multipagos.net/apoyo/API/pagos/add-payment/"

empresa= {'empresa':'CLARO'} 

#response = requests.post(url,params=empresa,auth=HTTPBasicAuth(usuario,contrasena ) )
response = requests.post(url2,params=empresa,auth=HTTPBasicAuth(usuario,contrasena ))

print(response.json())

