
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
dias = timedelta(days=7)
dias_ = timedelta(days=6)

f_inicio=(fecha-dias).strftime("%d-%m-%Y")
f_fin= ((fecha-dias)+dias_).strftime("%d-%m-%Y")

DATABASES = {
    'piki':{
        'NAME': 'piki',
        'USER': 'power_bi',
        'PASSWORD': 'PowerbI*456$44#',
        'HOST': '159.65.96.4',
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
SELECT 

pedidoid,facturaid,f_creacion,f_entrega,destinatario,
telefono,email,sucursal,departamento,afiliado,direccion_retiro,s_total,tax,total


 FROM
view_rpt_ventas where cast(f_entrega as date) between  timezone('America/Managua'::text, now())::date - 7 
and  timezone('America/Managua'::text, now())::date -1

and comercio in ('AmPm')
"""

detalle="""

SELECT 

pedidoid,facturaid,f_creacion,f_entrega,
destinatario,telefono,email,afiliado,sucursal,
precio,cantidad,
producto,categoria,
producto_descripcion,departamento


FROM
view_rpt_ventas_detalle 

where cast(f_entrega as date) between  timezone('America/Managua'::text, now())::date - 7 
and  timezone('America/Managua'::text, now())::date -1


and comercio in ('AmPm')
"""



# Create some Pandas dataframes from some data.
df_pedidos = pd.read_sql_query(pedidos,engine)

df_detalle = pd.read_sql_query(detalle,engine)



# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter(r'C:\Users\50576\Documents\Piki\AmPm.xlsx', engine='xlsxwriter')

# Write each dataframe to a different worksheet.
df_pedidos.to_excel(writer, index=False,sheet_name='ORDENES')
df_detalle.to_excel(writer, index=False,sheet_name='DETALLE')


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

email_list = [line.strip() for line in open(r'C:\Users\50576\Documents\Piki\email_ampm.txt')]

for to_addrs in email_list:
    smtp_user = 'keneth.morales@pikiapp.com'
    smtp_pass = 'of(820%D1jnA}G1]R,O-{'
    from_addr = 'keneth.morales@pikiapp.com'
    msg = MIMEMultipart()

    msg['Subject'] = "VENTAS PIKI "+ f_inicio +""" AL """+f_fin
    msg['From'] = from_addr
    msg['To'] = to_addrs
    
# Attach HTML to the email
    msg.attach(MIMEText("""
                         
                <b style="color: #24832B;">Buen dia </b>
                <br>      
                <br>           
                <b style="color: #777A77;font-size:12px;">Se adjunta detalle de ventas periodo:  """+f_inicio+""" al """+f_fin+""" </b>
                <br>  
                <br>
                <b style="color: #24832B;">Saludos</b>                                 """
                 ,'html'))

   

# Attach Cover Letter to the email
    msg.attach(load_file(r'C:\Users\50576\Documents\Piki\AmPm.xlsx','ventas_'+str(fecha)+'.xlsx'))

    try:
        sendmails(smtp_user, smtp_pass, 'keneth.morales@pikiapp.com', to_addrs, msg)
        print ("Email successfully sent to", to_addrs)
    except SMTPAuthenticationError:
        print ('SMTPAuthenticationError')
        print ("Email not sent to", to_addrs)
