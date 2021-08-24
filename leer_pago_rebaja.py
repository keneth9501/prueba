
import pandas as  pd
from datetime import datetime, timedelta
import shutil
import json
from requests.auth import HTTPDigestAuth
import requests
from requests.auth import HTTPBasicAuth



p_open=pd.read_excel(r'C:\Users\50576\Documents\Multipagos\open.xlsx',sheet_name='Pagos Open')
p_movil=pd.read_excel(r'C:\Users\50576\Documents\Multipagos\movil.xlsx',sheet_name='Pagos Movil')
p_promocion=pd.read_excel(r'C:\Users\50576\Documents\Multipagos\open.xlsx',sheet_name='Promocion')

p_open_nota=pd.read_excel(r'C:\Users\50576\Documents\Multipagos\open.xlsx',sheet_name='Creditos Open')
p_movil_nota=pd.read_excel(r'C:\Users\50576\Documents\Multipagos\movil.xlsx',sheet_name='Creditos Movil')


#########################################
p_movil_nota.rename(columns={
'CUSTOMER_ID':'codigo_cliente',
'REFERENCIA':'factura',
'Fecha_credito':'fecha_pago',
'Valor_Credito':'monto_cordoba'
},inplace=True)

p_movil_nota=p_movil_nota.assign(
    no_cupon=""
,no_fiscal=0,factura_interna=""
,referencia="",monto_dolar=0)


p_movil_nota=pd.DataFrame(p_movil_nota[
    ["codigo_cliente",
    "factura",
    "factura_interna",    
    'no_cupon',
    'no_fiscal',
    'referencia',
    'monto_cordoba',
    'monto_dolar',
    'fecha_pago']])
##########################################
p_open_nota.rename(columns={
'contrato':'codigo_cliente',
'factura_interna':'factura',
'numero_fiscal':'no_fiscal',
'Fecha_credito':'fecha_pago',
'Valor_Credito':'monto_cordoba'
},inplace=True)


p_open_nota=p_movil_nota.assign(
    no_cupon=""
,factura_interna=""
,referencia="",monto_dolar=0)


p_open_nota=pd.DataFrame(p_open_nota[
    ["codigo_cliente",
    "factura",
    "factura_interna",    
    'no_cupon',
    'no_fiscal',
    'referencia',
    'monto_cordoba',
    'monto_dolar',
    'fecha_pago']])

df=pd.concat([p_open_nota,p_movil_nota])

df['no_fiscal']=df['no_fiscal'].fillna(0)
df['no_fiscal']=df['no_fiscal'].astype(int)


p_open_=pd.DataFrame(p_open)
p_movil_=pd.DataFrame(p_movil)
p_promocion=p_promocion[p_promocion["Abono"]>1]
p_promocion["producto"]=1111


p_movil_['numero_fiscal']=0

p_movil_=pd.DataFrame(p_movil_[
    ["CO_ID",
    "CUSTOMER_ID",
    "CATEGORIA",
    "REFERENCIA",
    "numero_fiscal",
    "TELCONTACTO",
    "FECHA_ASIGNACION",
    "CICLO",
    "AÑO_FACTURA",
    "MESFACTURA",
    "CLIENTE",
    'TIPO_MORA2',
    'Fecha_Pago',
    'Abono']])

#renombrar columnas movil
p_movil_.rename(columns={'CO_ID':'producto',
'CUSTOMER_ID':'contrato',
'CATEGORIA':'Servicio',
'REFERENCIA':'factura_interna',
'numero_fiscal':'numero_fiscal',
'TELCONTACTO':'tel_contacto',
'FECHA_ASIGNACION':'fecha_asignacion',
'AÑO_FACTURA':'año',
'MESFACTURA':'mes',
'CICLO':'ciclo',
'CLIENTE':'suscriptor',
'TIPO_MORA2':'Tipo_Mora2',
'Fecha_Pago':'Fecha_Pago',
'Abono':'Abono'
},inplace=True)

# filtrar solo la  asignacion actual

# p_movil_.loc[( p_movil_['fecha_asignacion'].isin(['2019-10-25','2019-10-10'])) ]
# p_open_.loc[( p_open_['fecha_asignacion'].isin(['2019-10-25','2019-10-10'])) ]
# set fechas
#ok esta bien ahi dejalo ya lo veo .......

dfa=pd.concat([p_open_,p_movil_,p_promocion])

dfa['Fecha_Pago']=dfa['Fecha_Pago'].astype(int)
dfa['fecha_asignacion']=dfa['fecha_asignacion'].astype(int)

dfa['Fecha_Pago']=pd.to_datetime('1899-12-30')+pd.to_timedelta(dfa['Fecha_Pago'],'D')
dfa['fecha_asignacion']=pd.to_datetime('1899-12-30')+pd.to_timedelta(dfa['fecha_asignacion'],'D')

# dfa=pd.concat([p_open_,p_movil_]).groupby(by=['Fecha_Pago'],as_index=False)['Abono'].sum()
# dfa.name="Abono"
# dfa=dfa.reset_index()

dfa.to_excel(r'C:\Users\50576\Documents\Multipagos\Pagos_Cartera\072021.xlsx',engine='xlsxwriter',
index=False,sheet_name='pagos')

dfa.rename(columns={'contrato':'codigo_cliente',
                   'factura_interna':'factura',
                   'numero_fiscal':'no_fiscal',
                   'Fecha_Pago':'fecha_pago',
                   'Abono':'monto_cordoba',
                   },inplace=True)

dfa=dfa.assign(monto_dolar=0,no_cupon="",referencia="",factura_interna="",)


dfa=dfa[["codigo_cliente",
"factura",
"factura_interna",
"no_cupon",
"no_fiscal"
,"referencia",
"monto_cordoba",
"monto_dolar","fecha_pago"]]


usuario="keneth.morales"
contrasena="Alacran2020"


dfa['fecha_pago'] = dfa['fecha_pago'].dt.strftime('%Y-%m-%d')

dfa['no_fiscal']=dfa['no_fiscal'].fillna(0)
dfa['no_fiscal']=dfa['no_fiscal'].astype(int)

 

data_nota ={"empresa":"claro","tipo":"NOTA",
                         "data": df.to_dict(orient='records')
     }

data_pago ={"empresa":"claro","tipo":"PAGO",
                         "data": dfa.to_dict(orient='records')
     }



url = "http://sys.multipagos.net/apoyo/API/pagos/add-payment/"

# result = df.to_dict(orient='records')
# result={"empresa":"CLARO", "tipo":"nota",
#        "data": json.dumps(result,indent=4)}
# datos=json.loads(result)

datos_nota = json.dumps(data_nota, indent=0)
datos_pago = json.dumps(data_pago, indent=0)


with open('notas.txt', 'w') as outfile:
    outfile.write(datos_nota)
    outfile.write('\n')
    




# print(datos)

response = requests.post(url,
                        auth=HTTPBasicAuth(usuario,contrasena ), 
                        data=datos_pago)

                  


print(response.json())




response = requests.post(url,
                        auth=HTTPBasicAuth(usuario,contrasena ), 
                        data=datos_nota)                        


print(response.json())


# for i in range(len(p_open_)):
#   p_open_.loc[i,'Fecha_Pago']=pd.to_datetime('1899-12-30') + pd.to_timedelta(p_open_.loc[i,'Fecha_Pago'],'D')

  
  
  
   
  
  
  
  



  
 
	


