
import numpy as np
import smtplib,ssl 
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText 
from email.mime.base import MIMEBase 
from email import encoders
import pandas as pd
from sqlalchemy import create_engine

from datetime import datetime,timedelta
import email
import re   
from click import style
import html
from smtplib import SMTPAuthenticationError
from email.mime.image import MIMEImage
fecha=datetime.now()

dias = timedelta(days=1)

ayer=(fecha-dias).strftime("%d-%m-%Y")

DATABASES = {
    'piki':{
        'NAME': 'piki',
        'USER': 'power_bi',
        'PASSWORD': 'PowerbI*456$44#',
        'HOST': '162.243.160.180',
        'PORT': 5432,
    },
}

# choose the database to use
db = DATABASES['piki']

# construct an engine connection string
engine_string = "postgresql+psycopg2://{user}:{password}@{host}:{port}/{database}".format(
    user = db['USER'],
    password = db['PASSWORD'],
    host = db['HOST'],
    port = db['PORT'],
    database = db['NAME'],
)

# create sqlalchemy engine
engine = create_engine(engine_string)

pedidos="""
SELECT pedidoid,cast(f_entrega as date) as fecha, comercio,sucursal,afiliado,valor as delivery,s_total,tax,s_total+tax as total FROM
view_rpt_ventas where cast(f_entrega as date) = current_date - 1

and comercio in ('SUPER EXPRESS','Super Express Souvenirs')
"""

detalle="""

SELECT pedidoid,f_creacion,
f_entrega,destinatario
telefono,email,

direccion_retiro,direccion_entrega,
ubicacion_retiro,ubicacion_entrega,
sucursal,precio,cantidad,
CASE
WHEN upper(producto_descripcion) <> '' THEN
	upper(producto_descripcion)
WHEN upper(producto_descripcion) = '' THEN
	upper(producto)
ELSE
	'R'
END AS producto,categoria,producto_descripcion,comercio,afiliado,destinatario

 FROM
view_rpt_ventas_detalle where cast(f_entrega as date) = current_date - 1
and comercio in ('SUPER EXPRESS','Super Express Souvenirs')
"""

p_open=pd.read_excel(r'C:\Users\50576\Documents\Piki\descuento.xlsx',sheet_name='producto_descuento')
p_open_=pd.DataFrame(p_open)


# Create some Pandas dataframes from some data.
df_pedidos = pd.read_sql_query(pedidos,engine)

df_detalle = pd.read_sql_query(detalle,engine)

df_merge_difkey = pd.merge(df_detalle, p_open_, on='producto',how='left')


df_merge_difkey["descuento"] = (df_merge_difkey ["precio"] * df_merge_difkey ["cantidad"]) * df_merge_difkey ["%descuento"]

df_merge_difkey["precio sin iva"] = (df_merge_difkey ["precio"] * df_merge_difkey ["cantidad"])

df_merge_difkey["precio con descuento con iva"] = df_merge_difkey["precio sin iva"] - df_merge_difkey["descuento"]

df_merge_detalle_=pd.DataFrame(df_merge_difkey[['pedidoid', 
'f_creacion','f_entrega',
'destinatario','telefono',
'email','comercio',
'afiliado','direccion_retiro',
'direccion_entrega','ubicacion_retiro',
'ubicacion_entrega','sucursal',
'precio','cantidad',
'precio sin iva','%descuento',
'descuento','precio con descuento con iva',
'producto','categoria','producto_descripcion']])


# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter(r'C:\Users\50576\Documents\Piki\pandas_multiple.xlsx', engine='xlsxwriter')

# Write each dataframe to a different worksheet.
df_pedidos.to_excel(writer, index=False,sheet_name='pedidos')
df_merge_detalle_.to_excel(writer, index=False,sheet_name='detalle')


# Close the Pandas Excel writer and output the Excel file.
writer.save()


def load_file(file, file_name):

    read_file = open(file,'rb')
    attach = MIMEBase('application','vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    attach.set_payload(read_file.read()) 
    read_file.close()  
    encoders.encode_base64(attach) 
    attach.add_header('Content-Disposition', 'attachment', filename=file_name)
    return attach


def sendmails(username, password, from_addr, to_addrs, msg):
    
   
    
    context = ssl.create_default_context()
    smtp = smtplib.SMTP_SSL("box5725.bluehost.com",465,context=context)
    smtp.login(smtp_user,smtp_pass)
    smtp.sendmail(from_addr, to_addrs, msg.as_string())
    smtp.quit() 

email_list = [line.strip() for line in open(r'C:\Users\50576\Documents\Piki\email__.txt')]

for to_addrs in email_list:
    smtp_user = 'keneth.morales@pikiapp.com'
    smtp_pass = 'of(820%D1jnA}G1]R,O-{'
    from_addr = 'keneth.morales@pikiapp.com'
    msg = MIMEMultipart()

    msg['Subject'] = "VENTAS SUPER EXPRESS "+ayer
    msg['From'] = from_addr
    msg['To'] = to_addrs
    
# Attach HTML to the email
    msg.attach(MIMEText("""
                         
                <b>Buenos dias,
                <br>
                <br>
                Se adjunta detalle de ventas del dia """+ayer+"""
                <br>
                <br>
                
                <br>
                Saludos 
                                                 """
                 ,'html'))

   

# Attach Cover Letter to the email
    msg.attach(load_file(r'C:\Users\50576\Documents\Piki\pandas_multiple.xlsx','ventas_'+ayer+'.xlsx'))

    try:
        sendmails(smtp_user, smtp_pass, 'keneth.morales@pikiapp.com', to_addrs, msg)
        print ("Email successfully sent to", to_addrs)
    except SMTPAuthenticationError:
        print ('SMTPAuthenticationError')
        print ("Email not sent to", to_addrs)
