import pandas as  pd
from datetime import datetime, timedelta
import shutil
import os
import glob
import time
fecha="activacion_"+str(time.strftime("%d%m%y"))+".xlsx"

# path = r'C:\Users\50576\Desktop\mayo2020' 
# all_files = glob.glob(path + "/*.xlsx")

# li = []

# for filename in all_files:
#     df = pd.read_excel(filename, index_col=None, error_bad_lines=False)
#     li.append(df)

# frame = pd.concat(li, axis=0, ignore_index=True,sort=False)
# dfa=pd.DataFrame(frame)
# dfa=dfa.drop([0,1,2,3,4,5,6,7,8,9,10,11,12])
# dfa=dfa.drop(['Unnamed: 0', 'Unnamed: 1', 'Unnamed: 3','Unnamed: 9','Unnamed: 17','Unnamed: 18'], axis=1)
# dfa= dfa.dropna(axis=0, subset=['Unnamed: 2'])
# dfa.rename(columns={'Unnamed: 2':"transferencia",
# 'Unnamed: 4':"tipo",
# 'Unnamed: 5':"cliente",
# 'Unnamed: 6':"pos",
# 'Unnamed: 7':"service",
# 'Unnamed: 8':"pos_final",
# 'Unnamed: 13':"monto"
# },inplace=True)
# dfa.reset_index(inplace=True,drop=True)     

# dfa=pd.DataFrame(dfa[
# ["transferencia",
# "tipo",
# "cliente",
# "pos",  
# "service",
# "pos_final",
# "monto"]])
# dfa['fecha'] = dfa['transferencia'].str.slice(1, 7)
# dfa = dfa.drop(dfa[dfa['transferencia']=='ID de transacción'].index)

# dfa.to_excel(r'C:\Users\50576\Documents\TAE\Activacion\reactivaciones'+fecha+'.xlsx',engine='xlsxwriter',index=False,sheet_name='Pagos Open')
# print(dfa)




###############################################################


p_open=pd.read_excel(r'C:\Users\50576\Documents\TAE\VENTAS\c2sTransferChannelUserNew.xlsx')




p_open_=pd.DataFrame(p_open)
p_open_=p_open_.drop([0,1,2,3,4,5,6,7,8,9,10,11,12])
p_open_=p_open_.drop(['Unnamed: 0', 'Unnamed: 1', 'Unnamed: 3','Unnamed: 9','Unnamed: 17','Unnamed: 18'], axis=1)
p_open_= p_open_.dropna(axis=0, subset=['Unnamed: 2'])
p_open_.rename(columns={'Unnamed: 2':"transferencia",
'Unnamed: 4':"tipo",
'Unnamed: 5':"cliente",
'Unnamed: 6':"pos",
'Unnamed: 7':"service",
'Unnamed: 8':"pos_final",
'Unnamed: 13':"monto"
},inplace=True)
p_open_.reset_index(inplace=True,drop=True) 


p_open_=pd.DataFrame(p_open_[
    ["transferencia",
    "tipo",
    "cliente",
    "pos",  
    "service",
    "pos_final",
    "monto"]])

p_open_['fecha'] = p_open_['transferencia'].str.slice(1, 7)

p_open_.to_excel("C:/Users/50576/Documents/TAE/Activacion/%s" %fecha,
index=False,sheet_name='Pagos Open')


shutil.move(r'C:\Users\50576\Documents\TAE\VENTAS\c2sTransferChannelUserNew.xlsx', r'C:\Users\50576\Documents\TAE\RespaldoReactivacion\activacion.xlsx')

print(p_open_)
