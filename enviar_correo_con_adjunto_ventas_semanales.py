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

# read a table from database into pandas dataframe, replace "tablename" with your table name





query1="""select cast (f_entrega as date)fecha,count(pedidoid) as pedidos,sum(s_total) as ventas	
 from view_rpt_ventas where cast (f_entrega as date) between current_date - 8 and current_date -2
 and comercio in ('Don Pulpo') group by fecha
"""
query2=""" 
CREATE TEMPORARY TABLE comisiones (comercios varchar (250),tipo_pago varchar (250),condicion varchar (250), comision
VARCHAR(20));
insert into comisiones (comercios,tipo_pago,condicion,comision) values

('Asados All Carbon','CF','CREDITO','0.15'),
('Asados Don Pollo','CF','CREDITO','10'),
('Asados Los Robles','CF','CREDITO','0.15'),
('Baho Vilma','CF','CREDITO','0.15'),
('Be Vegan','CF','CREDITO','0.15'),
('BOTANICOS','CF','CREDITO','0.13'),
('Chocobania','CF','CREDITO','0.15'),
('Cocteleria El Bamboo','CF','CREDITO','0.15'),
('Comedor La Nani','CF','CREDITO','0.15'),
('D.CLEAN','CF','CREDITO','0.15'),
('DOGOOS','CF','CREDITO','0.15'),
('Don Lechon','CF','CREDITO','0.15'),
('Don Pulpo','CF','CREDITO','0.15'),
('Don Señor','CF','CREDITO','0.15'),
('El Calache Food Truck','CF','CREDITO','0.15'),
('El Eskimo','RET','CREDITO','0.15'),
('El Nuevo Muelle','CF','CREDITO','0.15'),
('El Patio de la Abuela','CF','CREDITO','0.13'),
('Freaking Frozen','CF','CREDITO','0.15'),
('Helados Frets','CF','CREDITO','0.15'),
('Hojarasca','CF','CREDITO','0.15'),
('LA BOQUERIA','CF','CREDITO','0.15'),
('La Crema Batida','RET','CREDITO','0.15'),
('La Gran Pizza','CF','CREDITO','0.15'),
('MadreSelva','CF','CREDITO','0.15'),
('Mariscos el Rey','RET','CREDITO','0.15'),
('Meson Español','RET','CREDITO','0.15'),
('Nicoya Bar & Grill','CF','CREDITO','0.15'),
('Patacones','CF','CREDITO','0.15'),
('Picanha Buffet Brasileiro','RET','CREDITO','0.15'),
('Pizza y Corre','CF','CREDITO','0.13'),
('Plata Con Plática','CF','CREDITO','0.13'),
('Restaurante Vegetariano ANANDA','RET','CREDITO','0.15'),
('Rispoli.s','CF','CREDITO','0.15'),
('RostiGO','CF','CREDITO','0.15'),
('Sand.s House','CF','CREDITO','0.15'),
('Shakes To Go','CF','CREDITO','0.13'),
('TGI FRIDAYS','RET','CREDITO','0.12'),
('Valenti.s Pizza','CF','CREDITO','0.15'),
('WAPI Batidos','CF','CREDITO','0.15'),
('Wasaby Sushi','CF','CREDITO','0.15'),
('Comedor El X Che','CF','CREDITO','0.15'),
('Xquisito','CF','CREDITO','0.15'),
('Carnes Shop','CF','CREDITO','0.15'),
('El Churraskito Grill','CF','CREDITO','0.15'),
('La Quequeria','CF','CREDITO','0.15'),
('Madre Pizza Napolitana','CF','CREDITO','0.15'),
('Tablitas Wings & Grill','CF','CREDITO','0.15'),
('Ristorante La Piazzeta','RET','CREDITO','0.15'),
('Delikatessen Quesos y Jamones','CF','CREDITO','0.12'),
('Mixes','CF','CREDITO','0.15'),
('Frontera Sur Beer & Wings','CF','CREDITO','0.15'),
('Xin Xing Restaurante','CF','CREDITO','0.15'),
('Cuba, mi vida','CF','CREDITO','0.15'),
('Antoja2 The Yard  ','CF','CREDITO','0.15'),
('Pollo Supremo','CF','CREDITO','0.15'),
('Simple Tech','CF','CREDITO','0.15'),
('Fritanga Nica','CF','CREDITO','0.15'),
('Entresabores','CF','CREDITO','0.13'),
('Mi Pancito','CF','CREDITO','0.15'),
('Stilettos','CF','CREDITO','0.15'),
('Lorraine Makeup','CF','CREDITO','0.15'),
('Da Raquelle','CF','CREDITO','0.15'),
('Alas Bravas','CF','CONTADO','0.13'),
('Pizza Fit','CF','CREDITO','0.13'),
('Toby.s Pizza','CF','CREDITO','0.13'),
('Floristeria Mil Flores','CF','CREDITO','0.15'),
('Sabor de Mi Tierra','CF','CREDITO','0.15'),
('Perro Bravo','CF','CREDITO','0.15'),
('Maika.i','CF','CREDITO','0.15'),
('Biibii','CF','CREDITO','0.15'),
('Momotombo Smoke','CF','CREDITO','0.15'),
('Gastro Las Colinas','CF','CREDITO','0.15'),
('Futec Industrial','CF','CREDITO','0.15'),
('PanPasión','CF','CREDITO','0.15'),
('Rock and Smoked','CF','CREDITO','0.15'),
('Café Orgánico JJ','CF','CREDITO','0.15'),
('Don Pancho','CF','CREDITO','0.13'),
('Restaurante Summer puerto Salvador Allende','RET','CREDITO','0.15'),
('Kofihub','CF','CREDITO','0.15'),
('La Capital','CF','CREDITO','0.15'),
('Mila.s Pancakes','CF','CREDITO','0.15'),
('Al Cilindro Nica','CF','CREDITO','0.14'),
('Amazonia','RET','CREDITO','0.14'),
('American Pizzería','CF','CREDITO','0.15'),
('Buffalo Wings','CF','CREDITO','0.1'),
('Chik Chak','RET','CREDITO','0.15'),
('La Estación','RET','CREDITO','0.25'),
('Pastelería Sampson','RET','CREDITO','0.14'),
('Santa Lucía','CF','CREDITO','0.14'),
('American Hamburguesas','CF','CREDITO','0.15'),
('Asados Doña Tania','CF','CONTADO','0.15'),
('Distribuidora La Universal','RET','CREDITO','0.15'),
('Restaurante El Gueguense','CF','CONTADO','0.12'),
('SUPER EXPRESS','CF','CONTADO','0.1'),
('Super Express Cervezas','CF','CREDITO','0.1'),
('Superfarmacias La Familiar','CF','CREDITO','0.1'),
('McDonald.s','CF','CREDITO','0.1'),
('Hooters','RET','CREDITO','0.13'),
('Farah Market Fit','CF','CREDITO','0.18'),
('Bottega La Pasta','CF','CREDITO','0.15'),
('Hard Rock Café','RET','CREDITO','0.14'),
('AJÚA','CF','CREDITO','0.15'),
('Bonnissimo Heladería','CF','CREDITO','0.15'),
('Leche Agria "El Ternero"','CF','CREDITO','0.15'),
('Hippos Galerias','RET','CREDITO','0.13'),
('Cafe Las Flores','RET','CREDITO','0.16'),
('Emporio Café, Bistro & Bar','RET','CREDITO','0.15'),
('Tacos Charros','RET','CONTADO','0.12'),
('Tip Top','CF','CREDITO','0.15'),
('Bubble Waffle Factory','RET','CREDITO','0.16'),
('Super Express Souvenirs','RET','CREDITO','0.1'),
('Don Ceviche','CF','CREDITO','0.15'),
('Woodys','RET','CREDITO','0.13'),
('Mesón Sur','CF','CREDITO','0.15'),
('El Parrillaje','CF','CREDITO','0.18'),
('Mancorp','CF','CREDITO','0.16'),
('Pikito.s','CF','CREDITO','0.15'),
('Amila’s Cocina Española','RET','CREDITO','0.18'),
('Sushi Itto','RET','CREDITO','0.16'),
('Stiletto’s Nicaragua','CF','CREDITO','0.15'),
('JCC Comercial','CF','CREDITO','0.15'),
('Como en Casa','CF','CREDITO','0.15'),
('Italianissimo','RET','CREDITO','0.16'),
('Asados El Patio','RET','CREDITO','0.16'),
('ACM Store Electronics','CF','CREDITO','0.18'),
('Bimbo','RET','CREDITO','0.15'),
('Helados Emanuel','CF','CREDITO','0.18'),
('Lorraine MakeUp','CF','CREDITO','0.15'),
('Café Las Flores','RET','CREDITO','0.16'),
('Nanah más que pupusas','CF','CREDITO','0.18'),
('Panadería La Fuente','RET','CREDITO','0.18'),
('La Carnicería','CF','CREDITO','0.15'),
('Super Herramientas','CF','CREDITO','0.14'),
('La Campana','RET','CREDITO','0.15'),
('Pizza Rustica','RET','CREDITO','0.15'),
('Don Parrillon','RET','CREDITO','0.13'),
('Huaraches','CF','CREDITO','0.18'),
('Pollos Narcy.s','CF','CREDITO','0.15'),
('Asados Momotombo','CF','CREDITO','0.18'),
('Chop Suey Internacional','RET','CREDITO','0.18'),
('El Bocadito','CF','CREDITO','0.15'),
('La Ventecita','CF','CONTADO','0.24'),
('Buffet Paladar','RET','CREDITO','0.24'),
('Asados Brisas y Más.','CF','CREDITO','0.18'),
('Nacatamales Doña Dora','CF','CREDITO','0.18'),
('Wings House','CF','CREDITO','0.18'),
('Nau Lounge','RET','CREDITO','0.15'),
('MyA Curiosidades','CF','CREDITO','0.18'),
('Rehab Bar  & Grill','RET','CREDITO','0.18'),
('Talu - Voces Vitales','CF','CREDITO','0.13'),
('Restaurante Summer','RET','CREDITO','0.18'),
('Arepas Venezolanas Maracay','CF','CREDITO','0.16'),
('Restaurantes Tip Top','CF','CREDITO','0.09'),
('Sabor de Mi Tierra Desayunos y más...','CF','CREDITO','0.18'),
('Asados Mary','CF','CREDITO','0.18'),
('Taqueria Los Tapatíos','RET','CREDITO','0.16'),
('Piki Mandados','CF','CONTADO','0.50'),

	
('SuplíMás Nicaragua','CF','CREDITO','0.15'),
('Al Sartén','RET','CREDITO','0.18'),
('Asados Tania','CF','CREDITO','0.18'),
('Choco Boom','CF','CREDITO','0.18'),
('Di que sí Coffe Shop','RET','CREDITO','0.18'),
('La Casa del CupCake de Sofía','CF','CREDITO','0.18'),
('La Colonia','RET','CONTADO','0.18'),
('La Tapiska','CF','CREDITO','0.18'),
('Michealitas','CF','CREDITO','0.18'),
('Pan del Banco','RET','CREDITO','0.18'),
('Prasad Foods','CF','CREDITO','0.18'),
('Típicos Las Flores','CF','CREDITO','0.18'),
('WingsBox','RET','CREDITO','0.18'),
('Bendición de Dios','CF','CREDITO','0.18'),
('Dulces delicias','CF','CREDITO','0.15'),
('La Casa de Fidel Ubau','CF','CREDITO','0.15'),
('La Gitana Sangria Artesanal','CF','CREDITO','0.18'),
('Pizza Tiffer','CF','CREDITO','0.18'),
('Riko sazon','RET','CREDITO','0.18'),
('Angelita Foods','RET','CREDITO','0.18')


;




select * from (

select * from (select 
ventas.comercios as comercio,

coalesce(sum(ventas.valor),'0') as delivery,

co.comision as comision,    
co.tipo_pago,
co.condicion,

 (count (ventas.envioid) ) as pedidos,
coalesce(sum(ventas.s_total),'0') total_venta,
coalesce(sum(ventas.tax),'0') iva,

CASE
   when ventas.comercios='Asados Don Pollo' then coalesce(count(ventas.envioid)* cast(comision as float))


    ELSE coalesce((sum(ventas.s_total)*cast(comision as float)),'0')
END AS comision_venta, 






CASE 
    WHEN co.condicion='CONTADO' THEN 0
    WHEN sum(ventas.tax) >  0 THEN coalesce((sum(ventas.s_total) - (sum(ventas.s_total)*cast(comision as float))) *1.15,'0')
    WHEN ventas.comercios= 'Asados Don Pollo'   THEN coalesce(sum(ventas.s_total),'0') - coalesce(count(ventas.envioid)* cast(comision as float))

    ELSE coalesce((sum(ventas.s_total) - (sum(ventas.s_total)*cast(comision as float))),'0')
END AS Neto_Pagar


  from(select *, coalesce(replace (comercio,'''','.'),'Piki Mandados') as comercios from view_rpt_ventas where cast (f_entrega as date) between 
timezone('America/Managua'::text, now())::date - 7 and timezone('America/Managua'::text, now())::date - 1) as ventas
left join (
select * from comisiones
) co

on ventas.comercios=co.comercios

where ventas.comercios not in('Wings House','Tip Top','SUPER 7') 
group by ventas.comercios,co.tipo_pago,co.comision,co.condicion


order by co.tipo_pago)

 as ventas1

union all

select * from 
(SELECT *,vb.total_venta * cast(vb.comision as float)   as comision_venta, 

vb.total_venta -(vb.total_venta * cast(vb.comision as float) ) as neto_pagar 

from  (select  replace (comercio,'''','.') || '-'|| sucursal  as comercios ,

sum(valor) as delivery,

case 
    WHEN replace (comercio,'''','.') || '-'|| sucursal  ~ '^Tip Top' then '0.15'
  
ELSE '0.18'

END as comision,

'CF' AS tipo_pago,
'CREDITO' AS condicion,
count (envioid) as pedidos,
sum(s_total) as total_venta,
sum(tax) as iva





 from view_rpt_ventas where cast (f_entrega as date) between 
timezone('America/Managua'::text, now())::date - 7 and timezone('America/Managua'::text, now())::date - 1 and comercio  in ('Wings House', 'Tip Top')

group by replace (comercio,'''','.'),sucursal)

as vb) as ventas2



UNION ALL

SELECT * FROM (SELECT *,vb.total_venta * cast(vb.comision as float)   as comision_venta, 

vb.total_venta -(vb.total_venta * cast(vb.comision as float) ) as neto_pagar 

from  (select  replace (comercio,'''','.') || '-'|| sucursal  as comercios ,

sum(valor) as delivery,

case 
    WHEN replace (comercio,'''','.') || '-'|| sucursal   ~ '^SUPER 7' then '0.15'
  
ELSE '0.18'

END as comision,

'CF' AS tipo_pago,
'CREDITO' AS condicion,
count (envioid) as pedidos,
sum(s_total) as total_venta,
sum(tax) as iva





 from view_rpt_ventas where cast (f_entrega as date) between 
timezone('America/Managua'::text, now())::date - 7 and timezone('America/Managua'::text, now())::date - 1 and comercio  in ('SUPER 7')

group by replace (comercio,'''','.'),sucursal)

as vb) AS ventas3



) as data_general

order by data_general.neto_pagar




"""

query3="""select coalesce (entidad_bancaria,'Efectivo') as Tipo_Pago, coalesce(sum (efectivo),'0') as efectivo,coalesce (sum (tarjeta),'0') as tarjeta 
from view_rpt_ventas
 where cast (f_entrega as date) between timezone('America/Managua'::text, now())::date - 7 and timezone('America/Managua'::text, now())::date - 1
 group by entidad_bancaria"""

query4="""select afiliado,sum( valor), count (afiliado) from view_rpt_ventas

where cast(f_entrega as date) between timezone('America/Managua'::text, now())::date - 7 and timezone('America/Managua'::text, now())::date - 1

group by afiliado"""


query5="""select afiliado,cast(f_entrega as date)  as f_entrega,sum( valor) from view_rpt_ventas

where cast(f_entrega as date) between timezone('America/Managua'::text, now())::date -7  and timezone('America/Managua'::text, now())::date - 1

group by afiliado,f_entrega"""




df = pd.read_sql_query(query1,engine)
df.to_excel('extraer_ventas_semanales_01.xlsx',engine='xlsxwriter',index=False,sheet_name='ventas')
df["pedidos"]=df["pedidos"].astype(int)
df["fecha"]=pd.to_datetime(df["fecha"]).dt.strftime('%d/%m/%Y')
df["comision 15%"]=df["ventas"] * 0.15
df["Neto Pagar"]= df["ventas"] - df["comision 15%"] 
df.loc['Total'] = df.sum(numeric_only=True)
df.loc['Total', 'fecha'] = "Total"




pd.options.display.float_format = '{:,2f}'.format
df2=pd.read_sql_query(query2,engine)
df2.sort_values(by=["neto_pagar"],inplace =True,ascending=False)
pedidos=df2['pedidos'].sum()
df2['pedidos']=df2['pedidos'].astype(str)
pd.set_option('display.width', 1000)
pd.options.display.float_format = '{:,.2f}'.format


df3 = pd.read_sql_query(query3,engine)
df3.loc['Total'] = df3.sum(numeric_only=True)
df3.loc['Total', 'tipo_pago'] = "Total"

df4 = pd.read_sql_query(query4,engine)
df4.loc['Total'] = df4.sum(numeric_only=True)
df4.loc['Total', 'afiliado'] = "Total"


df5 = pd.read_sql_query(query5,engine)
df5["f_entrega"]=pd.to_datetime(df5["f_entrega"]).dt.strftime('%d-%b')


df5=df5.pivot_table(index='afiliado',columns='f_entrega',values='sum',aggfunc=np.sum ).reset_index().rename_axis(None, axis=1)


df5.loc['Total'] = df5.sum(numeric_only=True)
df5.loc['Total', 'afiliado'] = "Total"

df5['TOTAL PAGO']=df5.sum(axis=1)

df5=pd.DataFrame(df5).fillna(0)


df3.rename(
    columns={'tipo_pago':'TIPO PAGO',
'efectivo':'EFECTIVO',
'tarjeta':'TARJETA'},
inplace=True
)

df5.rename(
    columns={'afiliado':'PIKER',

},
inplace=True
)

df4.rename(
    columns={'afiliado':'PIKER',
'sum':'PAGO',
'count':'PEDIDO'},
inplace=True
)



# # ## Configuramos separadores de miles y 2 decimales

# # # p=df.pivot_table(index='comercio', columns='f_entrega',values='', aggfunc=np.sum)
# # # #dar estilo a tabla html





df2=pd.DataFrame(df2)
df2.rename(columns={'comercio':'Comercio',
'delivery':'Delivery',
'comision':'%Comision',
'tipo_pago':'Tipo Pago',
'pedidos':'Pedidos',
'total_venta':'Venta Total',
'iva':'% 15iva',
'comision_venta':'Comision',
'neto_pagar':'Neto Pagar'

},inplace=True)





## fila total
df2.loc['Total'] = df2.sum(numeric_only=True)
df2.loc['Total', 'Comercio'] = "Total"
df2.loc['Total', 'Pedidos'] = pedidos.astype(str)
df2['%Comision'] = df2['%Comision'].fillna("-") 
df2['Tipo Pago'] = df2['Tipo Pago'].fillna("-") 
df2['Tipo Pago'].replace(['Cuota fija','Cuota Variable'],['CF','CV'],inplace=True)
df2.to_excel(r'C:\Users\50576\Documents\Piki\extraer_ventas_semanales.xlsx',engine='xlsxwriter',index=False,sheet_name='ventas')






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


def sendmail_detalle(username, password):
    smtp_user = 'keneth.morales@pikiapp.com'
    smtp_pass = 'of(820%D1jnA}G1]R,O-{'
    from_addr = 'keneth.morales@pikiapp.com'
    to_addrs  =  'conni193mn@gmail.com'


     
    
    
    notif = MIMEMultipart()

    notif['Subject'] = "Venta Semanal Comercios"
    notif['From'] = from_addr
    notif['To'] = to_addrs
    notif['Cc']="keneth.morales@metropolitanadistribucion.com"
# Attach HTML to the email
    notif.attach(MIMEText("""<b>Venta Semanal Comercios<br>
                Comercio: Don Pulpo<br>
                Periodo: """+f_inicio+""" al """+f_fin+""" <br></b>"""+html_style_basic(df)+"""<p><br>
                Para cualquier información, escribir a: maria.payan@pikiapp.com</p>""",'html'))
    msg.attach(load_file(r'C:\Users\50576\Documents\Piki\extraer_ventas_semanales_01.xlsx','ventas_'+f_inicio+"_"+f_fin+'.xlsx'))
           
    context = ssl.create_default_context()
    smtp = smtplib.SMTP_SSL("box5725.bluehost.com",465,context=context)
    smtp.login(smtp_user,smtp_pass)
    smtp.sendmail(from_addr, to_addrs, notif.as_string())
    smtp.quit() 
    
    print ("email sent to ",to_addrs)





def sendmails(username, password, from_addr, to_addrs, msg):
    
   
    
    context = ssl.create_default_context()
    smtp = smtplib.SMTP_SSL("box5725.bluehost.com",465,context=context)
    smtp.login(smtp_user,smtp_pass)
    smtp.sendmail(from_addr, to_addrs, msg.as_string())
    smtp.quit() 

email_list = [line.strip() for line in open(r'C:\Users\50576\Documents\Piki\email_ventas_semanales.txt')]

for to_addrs in email_list:
    smtp_user = 'keneth.morales@pikiapp.com'
    smtp_pass = 'of(820%D1jnA}G1]R,O-{'
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
                <b>DETALLE DE PAGO CUENTAS </b> 
                <br>
                """ +html_style_basic(df3)+"""
                <br>
                <b>DETALLE DE PAGO PIKER </b> 
                <br>
                """ +html_style_basic(df4)+"""
                <br>
                <b>DETALLE DE PAGO PIKER DIARIO </b> 
                <br>
                """ +html_style_basic(df5)+"""
                <br>
               Saludos """,'html'))
# Attach Cover Letter to the email
    msg.attach(load_file('C:/Users/50576/Documents/Piki/extraer_ventas_semanales.xlsx','ventas_.xlsx'))

    try:
        sendmails(smtp_user, smtp_pass, 'keneth.morales@pikiapp.com', to_addrs, msg)
        print ("Email successfully sent to", to_addrs)
        sendmail_detalle('keneth.morales@pikiapp.com','of(820%D1jnA}G1]R,O-{')
    except SMTPAuthenticationError:
        print ('SMTPAuthenticationError')
        print ("Email not sent to", to_addrs)




    



