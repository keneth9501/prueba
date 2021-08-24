import numpy as np
import urllib
import codecs
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

fecha=datetime.now()

dias = timedelta(days=7)
dias_ = timedelta(days=6)

 # Configuramos separadores de miles
pd.set_option('precision',0)

f_inicio=(fecha-dias).strftime("%d-%m-%Y")
f_fin= ((fecha-dias)+dias_).strftime("%d-%m-%Y")

print (f_inicio,f_fin)






# follows django database settings format, replace with your own settings
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

# read a table from database into pandas dataframe, replace "tablename" with your table name





query1="""SELECT cast(f_entrega as date) as fecha, case extract(dow from f_entrega)
            when 1 then 'Lunes'
            when 2 then 'Martes'
            when 3 then 'Miercoles'
            when 4 then 'Jueves'
            when 5 then 'Viernes'
            when 6 then 'Sabado'
            else 'Domingo'
       end as  Dia,count(envioid) as pedidos
, sum (s_total) as venta,SUM(tax) as iva,'13%' as comision,
sum (s_total)*0.13 as comisionpiki,

sum (s_total)-(sum (s_total)*0.13) as pago_comercio

FROM view_rpt_ventas
where comercio in ('Alas Bravas')

and cast (f_entrega as date) between 

timezone('America/Managua'::text, now())::date - 4 and (timezone('America/Managua'::text, now())::date - 1)

GROUP BY dia,fecha

order by fecha asc
"""









df=pd.read_sql_query(query1,engine)

df1=pd.DataFrame(df)

df1["pedidos"]=df["pedidos"].astype(int)
df1["fecha"]=pd.to_datetime(df["fecha"]).dt.strftime('%d/%m/%Y')
df1.loc['Total'] = df.sum(numeric_only=True)
df1.loc['Total', 'fecha'] = "Total"
pd.options.display.float_format = '{:,2f}'.format
pd.set_option('display.width', 1000)






df1.rename(
    columns={'fecha':'FECHA',
'dia':'DIA',
'pedidos':'PEDIDO',
'venta':'VENTA',
'iva':'IVA',
'comision':'COMISION %',
'comisionpiki':'COMISION PIKI',
'pago_comercio':'PAGO COMERCIO'




},
inplace=True
)



# # ## Configuramos separadores de miles y 2 decimales

# # # p=df.pivot_table(index='comercio', columns='f_entrega',values='', aggfunc=np.sum)
# # # #dar estilo a tabla html










## fila total





def html_style_basic(df,index=False):
    import pandas as pd
    x = df.to_html(index = index)
    x = x.replace('<table border="1" class="dataframe">','<table style="border-collapse: collapse; border-spacing: 0;last-child: color: #fff; width: 50%;">')
    x = x.replace('<th>','<th style="background-color: #29E768;font-size:15px;bgcolor:#E6E6FA;text-align: left; padding: 2px; border-left: 1px solid #0FD560;" align="left">')
    x = x.replace('<td>','<td style="text-align: left; padding: 1px; font-size:12px;border-left: 1px solid #B1B6BC; border-bottom: 1px solid #B1B6BC;border-right: 1px solid #B1B6BC;" align="left">')
    x = x.replace('<tr style="text-align: right;tr:last-child {	color: #fff;  background-color: #28B463;}">','<tr>')

    x = x.split()
    count = 2 
    index = 0
    for i in x:
        if '<tr>' in i:
            count+=1
            if count%2==0:
                x[index] = x[index].replace('<tr>','<tr style="background-color: #FFFFFF;" bgcolor="#FFFFFF">')
        index += 1
    return ' '.join(x)

#cargar archivo en ruta 

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

email_list = [line.strip() for line in open(r'C:\Users\50576\Documents\Piki\email_ventas_semanales.txt')]

for to_addrs in email_list:
    smtp_user = 'keneth.morales@pikiapp.com'
    smtp_pass = 'KmoraleS5482*12#'
    from_addr = 'keneth.morales@pikiapp.com'
    msg = MIMEMultipart()

    msg['Subject'] = "Venta Semanal Comercios "+f_inicio+" al"+" "+f_fin
    msg['From'] = from_addr
    msg['To'] = to_addrs
# Attach HTML to the email
    msg.attach(MIMEText("""
                <b>Venta Semanal Comercios
                <br>
                Periodo: """+f_inicio+""" al """+f_fin+""" 
                <br>
                """+html_style_basic(df)+"""
               
               Saludos """,'html'))
# Attach Cover Letter to the email
    msg.attach(load_file('C:/Users/50576/Documents/Piki/extraer_ventas_semanales.xlsx','ventas_.xlsx'))

    try:
        sendmails(smtp_user, smtp_pass, 'keneth.morales@pikiapp.com', to_addrs, msg)
        print ("Email successfully sent to", to_addrs)
        
    except SMTPAuthenticationError:
        print ('SMTPAuthenticationError')
        print ("Email not sent to", to_addrs)




    



