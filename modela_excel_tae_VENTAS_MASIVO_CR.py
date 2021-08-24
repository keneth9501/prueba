import pandas as  pd
from datetime import datetime, timedelta
import shutil
import os
import glob
import time
fecha="VENTA_MASIVO_"+str(time.strftime("%d%m%y"))+".xlsx"
p_open=pd.read_excel(r'C:\Users\50576\Documents\TAE\TAE CR\VENTAS_TEMP\C2cRetWidTransferChannelUser.xlsx')




p_open_=pd.DataFrame(p_open)




# p_open_=p_open_.drop([0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16])
# p_open_=p_open_.drop(['Unnamed: 0', 'Unnamed: 2', 'Unnamed: 6','Unnamed: 8','Unnamed: 10'], axis=1)
# p_open_= p_open_.dropna(axis=0, subset=['Unnamed: 1'])
# p_open_.rename(columns={'Unnamed: 4':"origen",
# 'Unnamed: 19':"monto_transferencia",
# 'Unnamed: 15':"fecha",
# 'Unnamed: 11':"transferencia",
# 'Unnamed: 9':"pos_destino",
# 'Unnamed: 5':"pos_origen",
# 'Unnamed: 7':"destino"
# },inplace=True)
# p_open_.reset_index(inplace=True,drop=True) 

# p_open_=pd.DataFrame(p_open_[
#     ["origen",
#     "pos_origen",
#     "destino",
#     "pos_destino",  
#     "transferencia",
#     "fecha",
#     "monto_transferencia"]])



p_open_.to_excel("C:/Users/50576/Documents/TAE/TAE CR/%s"%fecha,engine='xlsxwriter',
sheet_name='pagos')


# shutil.move(r'C:\Users\50576\Documents\TAE\TAE CR\VENTAS_TEMP\C2cRetWidTransferChannelUser.xlsx', r'C:\Users\50576\Documents\TAE\agosto2020.xlsx')

print(p_open_)
