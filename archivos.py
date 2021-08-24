import pandas as pd
import  glob

import numpy as np
# csv_files = glob.glob('C:/Users/50576/Documents/TAE/Activacion/*.xlsx')

# Mostrar el archivo csv_files, el cual es una lista de nombres
# list_data = []
  
# Escribimos un loop que irá a través de cada uno de los nombres de archivo a través de globbing y el resultado final será la lista dataframes
# for filename in csv_files:
#     data = pd.read_excel(filename)
#     list_data.append(data)
# Para chequear que todo está bien, mostramos la list_data por consola
# list_data
 
# df=pd.concat(list_data,ignore_index=True)
# df.to_csv('C:/Users/50576/Documents/TAE/Activacion/archivo.csv')

p_open=pd.read_csv( r'C:\Users\50576\Documents\archivo.csv')

p_open['fecha']=p_open['fecha'].astype(str)
p_open['fecha'] = p_open['fecha'].str.slice(0, 4)
p_open['fecha']=p_open['fecha'].astype(int)

p_open=p_open[p_open["fecha"]<=2104]

p_open=p_open[['pos_final','fecha']]

p_open.drop_duplicates('pos_final',inplace=True)

p_open=p_open.pivot_table(index='pos_final',columns='fecha',values='pos_final',aggfunc=np.count_nonzero ).reset_index().rename_axis(None, axis=1)

p_open=p_open.to_csv('C:/Users/50576/Documents/TAE/Activacion/archivo_pivot.csv')

print(p_open)
