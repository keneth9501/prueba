import pandas as pd 
import numpy as np

  
#df asignacion MOVIL 1
d1 = pd.ExcelFile(r'C:\Users\50576\Documents\Multipagos\12082021.xlsx') 
df_exc1 = d1.parse('movil') 
df_exc1['Servicio'] = 'MOVIL'+ ' '+ df_exc1['TIENE_FINANCIAMIENTO_EQUIPO']
dfa=pd.DataFrame(df_exc1[
    ['CUSTOMER_ID', 
    "Servicio",
    "DEPARTAMENTO",
    "MUNICIPIO",
    "SALDO",
    "CICLO",
    "AÑO_FACTURA",
    "MESFACTURA",
    "TIPO_CARTERA",
    "FECHA_ASI",
    "TIPO_MORA2"]])

dfa.rename(columns={'CUSTOMER_ID':"contrato",
"Servicio":"ServicioHmlg",
"DEPARTAMENTO":"departamento",
"MUNICIPIO":"localidad",    
"SALDO":"saldo_pendiente_factura",
"CICLO":"ciclo",
"AÑO_FACTURA":"año",
"MESFACTURA":"mes",
"TIPO_CARTERA":"tipo_cartera",
"FECHA_ASI":"fecha_asi",
"TIPO_MORA2":"TIPO_MORA_GRUPO"},inplace=True)

#df2 ASIGNACION OPEN 1
d2 = pd.ExcelFile(r'C:\Users\50576\Documents\Multipagos\12082021.xlsx') 
df_exc2 = d2.parse('open') 

financiamiento = df_exc2["FINAN_HOMOLOGADO"].copy()

df_exc2['ServicioHlmg'] = df_exc2['ServicioHlmg'].replace(regex='INDIVIDUAL', value='CASA CLARO')
    

df_exc2['ServicioHlmg']= df_exc2['ServicioHlmg']+ ' '+ financiamiento


dfb=pd.DataFrame(df_exc2[["contrato","ServicioHlmg","departamento","localidad",
"saldo_pendiente_factura","ciclo","año","mes","tipo_cartera","fecha_asi","TIPO_MORA_GRUPO"]])
###################################################################

dfb.rename(columns={'servicioHlmg':"ServicioHmlg"},inplace=True)


# ####################################################################################33
#df3 ASIGNACION MOVIL 2
# d3 = pd.ExcelFile(r'C:\Users\50576\Documents\Multipagos\28072021.xlsx') 
# df_exc3 = d3.parse('movil') 
# df_exc3['Servicio'] = 'MOVIL'+ ' '+ df_exc3['TIENE_FINANCIAMIENTO_EQUIPO']

# dfc=pd.DataFrame(df_exc3[
#       ['CUSTOMER_ID',
#        "Servicio",
#       "DEPARTAMENTO",
#       "MUNICIPIO",
#        "SALDO",
#        "CICLO",
#        "AÑO_FACTURA"
#        ,"MESFACTURA"
#       ,"TIPO_CARTERA"
#       ,"FECHA_ASI",
#       "TIPO_MORA2"]])
# dfc.rename(columns={'CUSTOMER_ID':"contrato",
# "Servicio":"ServicioHmlg",
# "DEPARTAMENTO":"departamento",
# "MUNICIPIO":"localidad",
# "SALDO":"saldo_pendiente_factura",
# "CICLO":"ciclo",
# "AÑO_FACTURA":"año",
# "MESFACTURA":"mes",
# "TIPO_CARTERA":"tipo_cartera",
# "FECHA_ASI":"fecha_asi",
# "TIPO_MORA2":"TIPO_MORA_GRUPO"},inplace=True)




# #df4 ASIGNACION OPEN 2
# d4 = pd.ExcelFile(r'C:\Users\50576\Documents\Multipagos\28072021.xlsx') 
# df_exc4 = d4.parse('open') 
# financiamiento = df_exc4["FINAN_HOMOLOGADO"].copy()

# df_exc4['ServicioHlmg'] = df_exc4['ServicioHlmg'].replace(regex='INDIVIDUAL', value='CASA CLARO')




# df_exc4['ServicioHlmg']= df_exc4['ServicioHlmg']+ ' '+ financiamiento

# dfd=pd.DataFrame(df_exc4[["contrato","ServicioHlmg","departamento","localidad",
# "saldo_pendiente_factura","ciclo","año","mes","tipo_cartera","fecha_asi","TIPO_MORA_GRUPO"]])

# dfd.rename(columns={'servicioHlmg':"ServicioHmlg"},inplace=True)
#############################################################################################

#UNION  PRIMERA ASIGNACION OPEN-MOVIL 
# df1_asig=pd.concat([dfc,dfd])

#UNION  SEGUNDA ASIGNACION OPEN-MOVIL 
df2_asig=pd.concat([dfa,dfb])


#UNION   ASIGNACION TOTAL  OPEN-MOVIL

df_union=pd.DataFrame(pd.concat([df2_asig]))

df_union_unique=pd.DataFrame(df_union[["contrato","ciclo","año","mes","tipo_cartera","fecha_asi","TIPO_MORA_GRUPO"]]) .drop_duplicates()



c=df_union.groupby(by=['contrato','ServicioHlmg','departamento','localidad'], as_index=False)['saldo_pendiente_factura'].sum()
c.name="saldo_total"
c=c.reset_index()

dfconsolidado= pd.merge(c,df_union_unique,on='contrato',how='inner')
dfconsolidado.drop_duplicates('contrato',keep="first",inplace=True)



#CREA ARCHIVO DE EXCEL CONSOLIDADO
df_union.to_excel(r'C:\Users\50576\Documents\Multipagos\Asignacion_Cartera\JUL202101.xlsx', engine='xlsxwriter',index=False)














#dfconsolidado.to_excel(r'C:\Users\50576\Documents\Multipagos\consolidado.xlsx', engine='xlsxwriter',index=False)


#df2_asig=pd.merge(dfd,dfc,left_on="contrato",right_on="CUSTOMER_ID",how='outer')
#df2_asig_=pd.DataFrame(df2_asig.iloc[:, 0:11])
#dfconsolida=pd.merge(df2_asig_,df1_asig_,on='contrato',how='outer')
#dfconsolida_=pd.DataFrame(dfconsolida.iloc[:, 0:11])
#df1_asig_.to_excel('consolidado2.xlsx', engine='xlsxwriter',index=False)









