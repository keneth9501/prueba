
from sqlalchemy import create_engine
import pandas as  pd
from datetime import datetime, timedelta
import shutil
import os
import glob
import time
fecha="activacion_"+str(time.strftime("%d%m%y"))+".xlsx"


DATABASES = {
    'apoyo':{
        'NAME': 'apoyo',
        'USER': 'kenth',
        'PASSWORD': 'K%8n)th#2020*',
        'HOST': 'db.multipagos.net',
        'PORT': 5432,
    },
}

# choose the database to use
db = DATABASES['apoyo']

# construct an engine connection string
engine_string = "postgresql+psycopg2://{user}:{password}@{host}:{port}/{database}".format(
    user = db['USER'],
    password = db['PASSWORD'],
    host = db['HOST'],
    port = db['PORT'],
    database = db['NAME'],
)


engine = create_engine(engine_string)

###############################################################

p_open=pd.read_excel(r'C:\Users\50576\Documents\Multipagos\Pago_Pretups\c2sTransferChannelUserNew.xlsx')
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

# df_merge.to_excel("C:/Users/50576/Documents/Multipagos/Pago_Pretups/%s" %fecha,
# index=False,sheet_name='Pagos Open')

df_merge.to_sql('apoyo_web_prectus', engine,if_exists='append',index=False)


shutil.move(r'C:\Users\50576\Documents\Multipagos\Pago_Pretups\c2sTransferChannelUserNew.xlsx', r'C:\Users\50576\Documents\Multipagos\Respaldo\activacion.xlsx')

print(p_open_)
