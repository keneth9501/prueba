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
pd.options.display.float_format = '{:,.0f}'.format
import xlsxwriter

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



detalle="""

with pedidos 

as(

select t1.*,t2.total_cervezas_pedido,

case 

when 	t2.total_cervezas_pedido>=12 then 'APLICA'

ELSE 'N/APLICA'

END AS APLICA








 from 

(select *,a2.cantidad_producto * a2.cervezas_contiene as total_cervezas



 from (




select a1.pedidoid,a1.fecha,a1.categoria,upper(a1.productos) as productos,afiliado,sum(cantidad_producto)as cantidad_producto, 

case

WHEN upper(productos)='CERVEZA TOÑA LATA 18X15' then 18
WHEN upper(productos)='CERVEZA TOÑA LITE BOTELLA 6 PACK' then 6
WHEN upper(productos)='CERVEZA TOÑA LITE BOTELLA 6PACK' then 6
WHEN upper(productos)='CERVEZA TOÑA LITE LATA UND' then 1
WHEN upper(productos)='CERVEZA VICTORIA FROST LATA 12X10' then 12
WHEN upper(productos)='TOÑA BOTELLA 350 ML DESECHABLE ( SIX PACK )' then 6
WHEN upper(productos)='TOÑA BOTELLA 350 ML DESECHABLE (SIX PACK)' then 6
WHEN upper(productos)='TOÑA BOTELLA 350 ML RETORNABLE' then 1
WHEN upper(productos)='TOÑA BOTELLA LITRO RETORNABLE' then 1
WHEN upper(productos)='TOÑA LATA 350 ML ( 12 PACK )' then 12
WHEN upper(productos)='TOÑA LATA 350 ML (12 PACK)' then 12
WHEN upper(productos)='TOÑA LATA 350 ML UND' then 1
WHEN upper(productos)='TOÑA LITE LATA 350 ML 6X5 SIX PACK' then 6
WHEN upper(productos)='VICTORIA BOTELLA 350 ML RETORNABLE' then 1
WHEN upper(productos)='VICTORIA BOTELLA LITRO RETORNABLE' then 1
WHEN upper(productos)='VICTORIA CLASICA BOTELLA DESECHABLE SIX PACK' then 6
WHEN upper(productos)='VICTORIA CLÁSICA BOTELLA DESECHABLE ( SIX PACK )' then 6
WHEN upper(productos)='VICTORIA CLÁSICA LATA 350 ML 15X12' then 15
WHEN upper(productos)='VICTORIA CLÁSICA LATA 350 ML 15X12 TWELVEPACK' then 15
WHEN upper(productos)='VICTORIA CLASICA LATA (SIX PACK)' then 6
WHEN upper(productos)='VICTORIA CLASICA LATA SIX PACK' then 6
WHEN upper(productos)='VICTORIA CLASICA LATA X 6' then 6
WHEN upper(productos)='VICTORIA FROST BOTELLA 350 ML DESECHABLE ( SIX PACK )' then 6
WHEN upper(productos)='VICTORIA FROST BOTELLA 350 ML DESECHABLE (SIX PACK)' then 6
WHEN upper(productos)='VICTORIA FROST BOTELLA 350 ML RETORNABLE' then 1
WHEN upper(productos)='VICTORIA FROST  BOTELLA LITRO RETORNABLE' then 1
WHEN upper(productos)='VICTORIA FROST LATA 350 ML (UNIDAD)' then 1
WHEN upper(productos)='VICTORIA LATA 350 ML (UNIDAD)' then 1
WHEN upper(productos)='VICTORIA SELECCION MAESTRO BOTELLA 350 ML DESECHABLE (SIX PACK)' then 6
WHEN upper(productos)='VICTORIA SELECCION MAESTRO LATA 350 ML (SIX PACK)' then 6
WHEN upper(productos)='BOTELLA 350 ML DESECHABLE (SIX PACK)' then 6
WHEN upper(productos)='LATA 350 ML  SIXPACK' then 6
WHEN upper(productos)='6X5 VICTORIA CLÁSICA BOTELLA DESECHABLE SUELTO' then 6
WHEN upper(productos)='6X5 VICTORIA FROST BOTELLA 350 ML DESECHABLE SUELTO' then 6
WHEN upper(productos)='VICTORIA CLÁSICA BOTELLA DESECHABLE SIX PACK' then 6

WHEN upper(productos)='TOÑA BOT 350ML RETORNABLE' then 1
WHEN upper(productos)='BEB010110 TOÑA BOTELLA 350 ML DESECHABLE (SIX PACK)' then 6
WHEN upper(productos)='VICTORIA BOT LITRO RETORNABLE' then 1
WHEN upper(productos)='TOÑA BOT LITRO RETORNABLE' then 1
WHEN upper(productos)='VICTORIA FROST BOT 350ML RETORNABLE' then 1
WHEN upper(productos)='VICTORIA FROST BOT LITRO RETORNABLE' then 1
WHEN upper(productos)='CERVEZA TOÑA LATA 16 OZ' then 1

WHEN upper(productos)='VICTORIA BOTELLA 350ML RETORNABLE' then 1
WHEN upper(productos)='VICTORIA SELECCION MAESTRO LATA 350 ML (UNIDAD)' then 1

WHEN upper(productos)='VICTORIA SELECCION MAESTRO BOTELLA 350 ML DESECHABLE (UNIDAD)' then 1
WHEN upper(productos)='TOÑA LITE BOT 350 ML RETORNABLE' then 1
WHEN upper(productos)='ENVASE LITRO PROY RETORNABLE AMBAR' then 1
WHEN upper(productos)='PROMO 6*5 CERVEZA TOÑA LITE LATA 350 ML' then 6

WHEN upper(productos)='VICTORIA FROST 16 ONZ 473 ML' then 1
WHEN upper(productos)='VICTORIA CLASICA LATA 473 ML 1X12' then 1

WHEN upper(productos)='COD. 212' then 1
WHEN upper(productos)='VICTORIA BOT 350ML RETORNABLE' then 1

WHEN upper(productos)='JC VICTORIA BOT LITRO RETORNABLE' then 1
WHEN upper(productos)='JC TOÑA BOT 350ML RETORNABLE' then 1


WHEN upper(productos)='MILLER LITE CERVEZA LATA SIX PACK 2124ML' then 6
WHEN upper(productos)='MILLER LITE CERVEZA BOTELLA DESECHABLE SIX PACK 1980ML' then 6
WHEN upper(productos)='CERVEZA TOÑA LITE LATA 6PACK' then 6
WHEN upper(productos)='VICTORIA FROST Lata 350 ml (12 pack)' then 12
WHEN upper(productos)='JC CERVEZA TOÑA LITE LATA 6PACK' then 12
WHEN upper(productos)='JC VICTORIA FROST Lata 350 ml (12 pack)' then 12
WHEN upper(productos)='VICTORIA FROST Lata 350 ml (12 pack)' then 12

WHEN upper(productos)='JC VICTORIA BOT 350ML RETORNABLE' then 1
WHEN upper(productos)='JC TOÑA BOT LITRO RETORNABLE' then 1
WHEN upper(productos)='JC VICTORIA LATA 350 ML (UNIDAD)' then 1
WHEN upper(productos)='VICTORIA FROST LATA 350 ML (12 PACK)' then 12

WHEN upper(productos)='JC VICTORIA FROST BOT 350ML RETORNABLE' then 12
WHEN upper(productos)='JC VICTORIA FROST LATA 350 ML (12 PACK)' then 1

WHEN upper(productos)='LATA 18X15' then 18
WHEN upper(productos)='LATA 350 ML 15X12 TWELVEPACK' then 15

WHEN upper(productos)='VICTORIA CLASICA BOTELLA 350ML RETORNABLE' then 1
WHEN upper(productos)='JC VICTORIA CLASIC BOT 350ML RETORNABLE' then 1

WHEN upper(productos)='VICTORIA CLASICA BOTELLA 350ML RETORNABLE' then 1
WHEN upper(productos)='JC VICTORIA CLASIC BOT 350ML RETORNABLE' then 1

WHEN upper(productos)='VICTORIA CLASICA BOTELLA LITRO RETORNABLE' then 1
WHEN upper(productos)='JC VICTORIA CLASIC LATA 350 ML (UNIDAD)' then 1

WHEN upper(productos)='JC VICTORIA FROST LATA 350 ML (UNIDAD)' then 1

else 0

end as cervezas_contiene

from 
(
SELECT
	pedidoid,
	cast(f_entrega as date) fecha,
	upper (categoria) as categoria,
	CASE
WHEN producto_descripcion <> '' THEN
	producto_descripcion
WHEN producto_descripcion = '' THEN
	producto
ELSE
	'R'
END AS productos,
afiliado,

cantidad as cantidad_producto

FROM
	view_rpt_ventas_detalle
where
 comercio IN (
	'SUPER EXPRESS',
	'Super Express Cervezas'
)
AND upper (categoria) IN ('CERVEZA', 'RETORNABLES','PROMOCIÓN','JUEVES CERVECERO','NACIONALES'))

 as a1

group by a1.pedidoid,a1.fecha,a1.categoria,upper(a1.productos),afiliado



) as a2

 ) as t1

left join (


select a2.pedidoid,sum(a2.cantidad_producto * a2.cervezas_contiene) as total_cervezas_pedido



 from (




select a1.pedidoid,a1.fecha,a1.categoria,upper(a1.productos) as productos,afiliado,sum(cantidad_producto)as cantidad_producto, 

case

WHEN upper(productos)='CERVEZA TOÑA LATA 18X15' then 18
WHEN upper(productos)='CERVEZA TOÑA LITE BOTELLA 6 PACK' then 6
WHEN upper(productos)='CERVEZA TOÑA LITE BOTELLA 6PACK' then 6
WHEN upper(productos)='CERVEZA TOÑA LITE LATA UND' then 1
WHEN upper(productos)='CERVEZA VICTORIA FROST LATA 12X10' then 12
WHEN upper(productos)='TOÑA BOTELLA 350 ML DESECHABLE ( SIX PACK )' then 6
WHEN upper(productos)='TOÑA BOTELLA 350 ML DESECHABLE (SIX PACK)' then 6
WHEN upper(productos)='TOÑA BOTELLA 350 ML RETORNABLE' then 1
WHEN upper(productos)='TOÑA BOTELLA LITRO RETORNABLE' then 1
WHEN upper(productos)='TOÑA LATA 350 ML ( 12 PACK )' then 12
WHEN upper(productos)='TOÑA LATA 350 ML (12 PACK)' then 12
WHEN upper(productos)='TOÑA LATA 350 ML UND' then 1
WHEN upper(productos)='TOÑA LITE LATA 350 ML 6X5 SIX PACK' then 6
WHEN upper(productos)='VICTORIA BOTELLA 350 ML RETORNABLE' then 1
WHEN upper(productos)='VICTORIA BOTELLA LITRO RETORNABLE' then 1
WHEN upper(productos)='VICTORIA CLASICA BOTELLA DESECHABLE SIX PACK' then 6
WHEN upper(productos)='VICTORIA CLÁSICA BOTELLA DESECHABLE ( SIX PACK )' then 6
WHEN upper(productos)='VICTORIA CLÁSICA LATA 350 ML 15X12' then 15
WHEN upper(productos)='VICTORIA CLÁSICA LATA 350 ML 15X12 TWELVEPACK' then 15
WHEN upper(productos)='VICTORIA CLASICA LATA (SIX PACK)' then 6
WHEN upper(productos)='VICTORIA CLASICA LATA SIX PACK' then 6
WHEN upper(productos)='VICTORIA CLASICA LATA X 6' then 6
WHEN upper(productos)='VICTORIA FROST BOTELLA 350 ML DESECHABLE ( SIX PACK )' then 6
WHEN upper(productos)='VICTORIA FROST BOTELLA 350 ML DESECHABLE (SIX PACK)' then 6
WHEN upper(productos)='VICTORIA FROST BOTELLA 350 ML RETORNABLE' then 1
WHEN upper(productos)='VICTORIA FROST  BOTELLA LITRO RETORNABLE' then 1
WHEN upper(productos)='VICTORIA FROST LATA 350 ML (UNIDAD)' then 1
WHEN upper(productos)='VICTORIA LATA 350 ML (UNIDAD)' then 1
WHEN upper(productos)='VICTORIA SELECCION MAESTRO BOTELLA 350 ML DESECHABLE (SIX PACK)' then 6
WHEN upper(productos)='VICTORIA SELECCION MAESTRO LATA 350 ML (SIX PACK)' then 6
WHEN upper(productos)='BOTELLA 350 ML DESECHABLE (SIX PACK)' then 6
WHEN upper(productos)='LATA 350 ML  SIXPACK' then 6
WHEN upper(productos)='6X5 VICTORIA CLÁSICA BOTELLA DESECHABLE SUELTO' then 6
WHEN upper(productos)='6X5 VICTORIA FROST BOTELLA 350 ML DESECHABLE SUELTO' then 6
WHEN upper(productos)='VICTORIA CLÁSICA BOTELLA DESECHABLE SIX PACK' then 6


WHEN upper(productos)='TOÑA BOT 350ML RETORNABLE' then 1
WHEN upper(productos)='BEB010110 TOÑA BOTELLA 350 ML DESECHABLE (SIX PACK)' then 6
WHEN upper(productos)='VICTORIA BOT LITRO RETORNABLE' then 1
WHEN upper(productos)='TOÑA BOT LITRO RETORNABLE' then 1
WHEN upper(productos)='VICTORIA FROST BOT 350ML RETORNABLE' then 1
WHEN upper(productos)='VICTORIA FROST BOT LITRO RETORNABLE' then 1
WHEN upper(productos)='CERVEZA TOÑA LATA 16 OZ' then 1

WHEN upper(productos)='VICTORIA BOTELLA 350ML RETORNABLE' then 1
WHEN upper(productos)='VICTORIA SELECCION MAESTRO LATA 350 ML (UNIDAD)' then 1

WHEN upper(productos)='VICTORIA SELECCION MAESTRO BOTELLA 350 ML DESECHABLE (UNIDAD)' then 1
WHEN upper(productos)='TOÑA LITE BOT 350 ML RETORNABLE' then 1
WHEN upper(productos)='ENVASE LITRO PROY RETORNABLE AMBAR' then 1
WHEN upper(productos)='PROMO 6*5 CERVEZA TOÑA LITE LATA 350 ML' then 6

WHEN upper(productos)='VICTORIA FROST 16 ONZ 473 ML' then 1
WHEN upper(productos)='VICTORIA CLASICA LATA 473 ML 1X12' then 1

WHEN upper(productos)='COD. 212' then 1
WHEN upper(productos)='VICTORIA BOT 350ML RETORNABLE' then 1

WHEN upper(productos)='JC VICTORIA BOT LITRO RETORNABLE' then 1
WHEN upper(productos)='JC TOÑA BOT 350ML RETORNABLE' then 1




WHEN upper(productos)='MILLER LITE CERVEZA LATA SIX PACK 2124ML' then 6
WHEN upper(productos)='MILLER LITE CERVEZA BOTELLA DESECHABLE SIX PACK 1980ML' then 6
WHEN upper(productos)='CERVEZA TOÑA LITE LATA 6PACK' then 6
WHEN upper(productos)='VICTORIA FROST Lata 350 ml (12 pack)' then 12
WHEN upper(productos)='JC CERVEZA TOÑA LITE LATA 6PACK' then 12
WHEN upper(productos)='JC VICTORIA FROST Lata 350 ml (12 pack)' then 12
WHEN upper(productos)='VICTORIA FROST Lata 350 ml (12 pack)' then 12

WHEN upper(productos)='JC VICTORIA BOT 350ML RETORNABLE' then 1
WHEN upper(productos)='JC TOÑA BOT LITRO RETORNABLE' then 1
WHEN upper(productos)='JC VICTORIA LATA 350 ML (UNIDAD)' then 1
WHEN upper(productos)='VICTORIA FROST LATA 350 ML (12 PACK)' then 12

WHEN upper(productos)='JC VICTORIA FROST BOT 350ML RETORNABLE' then 12
WHEN upper(productos)='JC VICTORIA FROST LATA 350 ML (12 PACK)' then 1

WHEN upper(productos)='LATA 18X15' then 18
WHEN upper(productos)='LATA 350 ML 15X12 TWELVEPACK' then 15

WHEN upper(productos)='VICTORIA CLASICA BOTELLA 350ML RETORNABLE' then 1
WHEN upper(productos)='JC VICTORIA CLASIC BOT 350ML RETORNABLE' then 1

WHEN upper(productos)='VICTORIA CLASICA BOTELLA LITRO RETORNABLE' then 1
WHEN upper(productos)='JC VICTORIA CLASIC LATA 350 ML (UNIDAD)' then 1

WHEN upper(productos)='JC VICTORIA FROST LATA 350 ML (UNIDAD)' then 1


else 0

end as cervezas_contiene

from 
(
SELECT
	pedidoid,
	cast(f_entrega as date) fecha,
	upper (categoria) as categoria,
	CASE
WHEN producto_descripcion <> '' THEN
	producto_descripcion
WHEN producto_descripcion = '' THEN
	producto
ELSE
	'R'
END AS productos,
afiliado,

cantidad as cantidad_producto

FROM
	view_rpt_ventas_detalle
where
 comercio IN (
	'SUPER EXPRESS',
	'Super Express Cervezas'
)
AND upper (categoria) IN ('CERVEZA', 'RETORNABLES','PROMOCIÓN','JUEVES CERVECERO','NACIONALES'))

as a1

group by a1.pedidoid,a1.fecha,a1.categoria,upper(a1.productos),afiliado



) as a2

group by a2.pedidoid
)

t2

on t1.pedidoid=t2.pedidoid

)

select * from pedidos
where cast (fecha as date) between
timezone('America/Managua'::text, now())::date -7  and timezone('America/Managua'::text, now())::date - 1

"""
df_pedidos_detalle = pd.read_sql_query(detalle,engine)


df_pedidos = df_pedidos_detalle[['pedidoid', 'fecha','categoria','total_cervezas_pedido','aplica']].drop_duplicates()
df_pedidos["total_cervezas_pedido"]=df_pedidos["total_cervezas_pedido"].astype(int)
df_pedidos_detalle["fecha"]=df_pedidos_detalle["fecha"].astype(str)
df_pedidos["fecha"]=df_pedidos["fecha"].astype(str)


p = df_pedidos.groupby(['pedidoid']).size().reset_index(name='counts')

df_count= pd.DataFrame(p)

df_merge= pd.merge(df_pedidos, df_count, on='pedidoid',how='left')
# df_merge[['pedidoid', 'categoria',"0"]].loc[df_merge["0"] == 2]
# df_merging=df_merge.rename(columns = {list(df_merge)[5]:'cuenta'}, inplace=True)

###df2 = df_merge.drop(df_merge[(df_merge["categoria"].isin(["Cerveza","Promoción"])) & (df_merge["counts"] == 2)].index)

df_merge.sort_values(by=["categoria"],inplace =True,ascending=False)
df_merge.drop_duplicates('pedidoid',keep="first",inplace=True)

df_tabla=df_merge.pivot_table(index='categoria',columns='aplica',values='pedidoid',aggfunc=np.count_nonzero ).reset_index().rename_axis(None, axis=1)
df_tabla.loc['TOTAL'] = df_tabla.sum(numeric_only=True).astype(int)

df_tabla['TOTAL']=df_tabla.sum(numeric_only=True,axis=1).astype(int)
df_tabla.loc['TOTAL', 'categoria'] = "TOTAL"

df_tabla=pd.DataFrame(df_tabla).fillna(0)



df_tabla.rename(columns={'categoria':'CATEGORIA'

},inplace=True)



# Write each dataframe to a different worksheet.

columnas=['pedidoid','fecha','categoria','total_cervezas_pedido','aplica']



data = pd.DataFrame({'a':[1,2,3,4,5],
                   'b':[2,3,4,5,6],
                   'c':[3,4,5,6,7],
                   'd':[4,5,6,7,8],
                   'e':[5,6,7,8,9]})


dfdetalle=df_pedidos_detalle.values.tolist()				   
df=df_merge.values.tolist()
df_resumen=df_tabla.values.tolist()



caption = 'PEDIDOS APLICA DELIVERY SUPER EXPRESS'



workbook = xlsxwriter.Workbook(r"C:\Users\50576\Documents\Piki\aplica-delivery.xlsx")
bold = workbook.add_format({'bold': True})
worksheet1 = workbook.add_worksheet("pedidos")
worksheet2= workbook.add_worksheet("detalle")
worksheet1.write('A1', caption,bold)




worksheet2.add_table('A1:J700', {'data':dfdetalle,
'style': 'Table Style Medium 2','header_row': True,
'columns': [{'header': 'PEDIDOID'},
            {'header': 'FECHA'},
            {'header': 'CATEGORIA'},
            {'header': 'PRODUCTO'},
            {'header': 'AFILIADO'},
            {'header': 'CANTIDAD PRODUCTO'},  
                                      
            {'header': 'CERVEZAS CONTIENE'},
            {'header': 'TOTAL CERVEZAS'},
            {'header': 'TOTAL CERVEZAS PEDIDO'},
            {'header': 'APLICA'},
]
})


worksheet1.add_table('A3:E600', {'data':df,
'style': 'Table Style Medium 2','header_row': True,
'columns': [{'header': 'PEDIDOID'},
            {'header': 'FECHA'},
            {'header': 'CATEGORIA'},
            {'header': 'TOTAL CERVEZA'},
            {'header': 'APLICA'},
                                          ]


})


worksheet1.add_table('H1:K5', {'first_column': True,'last_column': True,'autofilter': False,'data':df_resumen,
'style': 'Table Style Medium 4','header_row': True,
'columns': [{'header': 'CATEGORIA'},
            {'header': 'APLICA'},
            {'header': 'NO APLICA'},
            {'header': 'TOTAL'},
			
			

                                                      ]


})






worksheet1.hide_gridlines(2)
worksheet2.hide_gridlines(2)


workbook.close()

def html_style_basic(df,index=False):
    import pandas as pd
    x = df.to_html(index = index)
    x = x.replace('<table border="1" class="dataframe">','<table style="border-collapse: collapse; border-spacing: 0;last-child: color: #E4EFF6; width: 40%;">')
    x = x.replace('<th>','<th style="background-color: #09AA27;color: #FCFBFB;font-size:11px;bgcolor:#E6E6FA;text-align: left; padding: 3px; border-left: 2px solid #B0B3B6;" align="left">')
    x = x.replace('<td>','<td style="text-align: left; padding: 1px; color: #364238;font-size:9px;border-left: 1px solid #B1B6BC; border-bottom: 1px solid #B1B6BC;border-right: 1px solid #B1B6BC;" align="left">')
    x = x.replace('<tr style="text-align: right;last-child {    background: blue;}">','<tr>')

    x = x.split()
    count = 2 
    index = 0
    for i in x:
        if '<tr>' in i:
            count+=1
            if count%2==0:
                x[index] = x[index].replace('<tr>','<tr style="background-color: #D3FADA;" bgcolor="#D3FADA">')
        index += 1
    return ' '.join(x)

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
    smtp_pass = 'of(820%D1jnA}G1]R,O-{'
    from_addr = 'keneth.morales@pikiapp.com'
    msg = MIMEMultipart()

    msg['Subject'] = "#ORDENES APLICA DELIVERY"
    msg['From'] = from_addr
    msg['To'] = to_addrs
    
# Attach HTML to the email
    msg.attach(MIMEText("""
                         
                <b style="color: #EE7313;">COMERCIO:SUPER EXPRESS </b>
                <br>                
                <b style="color: #758978;font-size:10px;">PERIODO: """+f_inicio+""" al """+f_fin+""" </b>
                <b>
                """+html_style_basic(df_tabla)+"""
                                                 """
                 ,'html'))

   

# Attach Cover Letter to the email
    msg.attach(load_file(r'C:\Users\50576\Documents\Piki\aplica-delivery.xlsx','PEDIDOS.xlsx'))

    try:
        sendmails(smtp_user, smtp_pass, 'keneth.morales@pikiapp.com', to_addrs, msg)
        print ("Email successfully sent to", to_addrs)
    except SMTPAuthenticationError:
        print ('SMTPAuthenticationError')
        print ("Email not sent to", to_addrs)






















# Close the Pandas Excel writer and output the Excel file.









