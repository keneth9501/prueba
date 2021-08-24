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
select 
a.pedidoid,
a.comercio,
a.piker,
a.f_creacion::text,
a.f_asignacion::text,
a.f_recepcion::text,
a.f_entrega::text,
a.hora_creacion::text,
a.hora_asignacion::text,
a.hora_recepcion::text,
a.hora_entrega::text,
extract(min from date_trunc('min',hora_entrega - hora_creacion + interval '25 second')) as tiempo_total,
extract(min from date_trunc('min',hora_entrega - hora_recepcion + interval '25 second')) as tiempo_entrega,
extract (min from date_trunc('min',hora_recepcion - hora_creacion + interval '25 second')) as tiempo_preparacion
from (
SELECT
	pedidoid,
  comercio,
  afiliado as piker,
CAST (f_creacion AS DATE) AS f_creacion,
CAST (fecha_asignacion AS DATE) AS f_asignacion,
CAST (fecha_recepcion AS DATE) AS f_recepcion,
CAST (f_entrega AS DATE) AS f_entrega,
	timezone('GMT'::text, f_creacion)::time AS hora_creacion,
	
	timezone('GMT'::text, fecha_asignacion)::time AS hora_asignacion,

	timezone('GMT'::text, fecha_recepcion)::time AS hora_recepcion,
	
  timezone('GMT'::text, f_entrega)::time AS hora_entrega


FROM
	view_rpt_ventas
where cast(f_entrega as date) = timezone('America/Managua'::text, now())::date -1
) as a

"""
df_pedidos_detalle = pd.read_sql_query(detalle,engine)


# df_pedidos = df_pedidos_detalle[['pedidoid', 'fecha','categoria','total_cervezas_pedido','aplica']].drop_duplicates()
# df_pedidos["total_cervezas_pedido"]=df_pedidos["total_cervezas_pedido"].astype(int)
# df_pedidos_detalle["fecha"]=df_pedidos_detalle["fecha"].astype(str)
# df_pedidos["fecha"]=df_pedidos["fecha"].astype(str)


# p = df_pedidos.groupby(['pedidoid']).size().reset_index(name='counts')

# df_count= pd.DataFrame(p)

# df_merge= pd.merge(df_pedidos, df_count, on='pedidoid',how='left')
# df_merge[['pedidoid', 'categoria',"0"]].loc[df_merge["0"] == 2]
# df_merging=df_merge.rename(columns = {list(df_merge)[5]:'cuenta'}, inplace=True)

###df2 = df_merge.drop(df_merge[(df_merge["categoria"].isin(["Cerveza","Promoción"])) & (df_merge["counts"] == 2)].index)

# df_merge.sort_values(by=["categoria"],inplace =True,ascending=False)
# df_merge.drop_duplicates('pedidoid',keep="first",inplace=True)

# df_tabla=df_merge.pivot_table(index='categoria',columns='aplica',values='pedidoid',aggfunc=np.count_nonzero ).reset_index().rename_axis(None, axis=1)
# df_tabla.loc['TOTAL'] = df_tabla.sum(numeric_only=True).astype(int)

# df_tabla['TOTAL']=df_tabla.sum(numeric_only=True,axis=1).astype(int)
# df_tabla.loc['TOTAL', 'categoria'] = "TOTAL"

# df_tabla=pd.DataFrame(df_tabla).fillna(0)



# df_tabla.rename(columns={'categoria':'CATEGORIA'

# },inplace=True)



# # Write each dataframe to a different worksheet.

# columnas=['pedidoid','fecha','categoria','total_cervezas_pedido','aplica']



# data = pd.DataFrame({'a':[1,2,3,4,5],
#                    'b':[2,3,4,5,6],
#                    'c':[3,4,5,6,7],
#                    'd':[4,5,6,7,8],
#                    'e':[5,6,7,8,9]})


dfdetalle=df_pedidos_detalle.values.tolist()				   
# df=df_merge.values.tolist()
# df_resumen=df_tabla.values.tolist()



# caption = 'PEDIDOS APLICA DELIVERY SUPER EXPRESS'



workbook = xlsxwriter.Workbook(r"C:\Users\50576\Documents\Piki\tiempos-entrega.xlsx")
bold = workbook.add_format({'bold': True})
worksheet2 = workbook.add_worksheet("pedidos")
# worksheet2= workbook.add_worksheet("detalle")
# worksheet1.write('A1', caption,bold)




worksheet2.add_table('A1:N700', {'data':dfdetalle,
'style': 'Table Style Medium 2','header_row': True,
'columns': [{'header': 'PEDIDOID'},
            {'header': 'COMERCIO'},
            {'header': 'PIKER'},
            {'header': 'F_CREACION'},
            {'header': 'F_ASIGNACION'},
            {'header': 'F_RECEPCION'},  
                                      
            {'header': 'F_ENTREGA'},
            {'header': 'HORA_CREACION'},
            {'header': 'HORA_ASIGNACION'},
            {'header': 'HORA_RECEPCION'},
             {'header': 'HORA_ENTREGA'},
             {'header': 'TIEMPO_TOTAL'},
             {'header': 'TIEMPO_ENTREGA'},
             {'header': 'TIMEPO_PREPARACION'}


]
})


# worksheet1.add_table('A3:E600', {'data':df,
# 'style': 'Table Style Medium 2','header_row': True,
# 'columns': [{'header': 'PEDIDOID'},
#             {'header': 'FECHA'},
#             {'header': 'CATEGORIA'},
#             {'header': 'TOTAL CERVEZA'},
#             {'header': 'APLICA'},
#                                           ]


# })


# worksheet1.add_table('H1:K5', {'first_column': True,'last_column': True,'autofilter': False,'data':df_resumen,
# 'style': 'Table Style Medium 4','header_row': True,
# 'columns': [{'header': 'CATEGORIA'},
#             {'header': 'APLICA'},
#             {'header': 'NO APLICA'},
#             {'header': 'TOTAL'},
			
			

#                                                       ]


# })






# worksheet1.hide_gridlines(2)
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

email_list = [line.strip() for line in open(r'C:\Users\50576\Documents\Piki\email_tiempo_entrega.txt')]

for to_addrs in email_list:
    smtp_user = 'keneth.morales@pikiapp.com'
    smtp_pass = 'of(820%D1jnA}G1]R,O-{'
    from_addr = 'keneth.morales@pikiapp.com'
    msg = MIMEMultipart()

    msg['Subject'] = "REPORTE ORDENES - TIEMPO DE ENTREGA"
    msg['From'] = from_addr
    msg['To'] = to_addrs
    
# Attach HTML to the email
    msg.attach(MIMEText("""
                         
                <b style="color: #EE7313;">TIEMPOS DE ENTREGA PEDIDOS </b>
                <br>                
                <b style="color: #758978;font-size:10px;">PERIODO: """+ayer+""" </b>
                <b>
                """
                 ,'html'))

   

# Attach Cover Letter to the email
    msg.attach(load_file(r'C:\Users\50576\Documents\Piki\tiempos-entrega.xlsx','PEDIDOS.xlsx'))

    try:
        sendmails(smtp_user, smtp_pass, 'keneth.morales@pikiapp.com', to_addrs, msg)
        print ("Email successfully sent to", to_addrs)
    except SMTPAuthenticationError:
        print ('SMTPAuthenticationError')
        print ("Email not sent to", to_addrs)






















# Close the Pandas Excel writer and output the Excel file.









