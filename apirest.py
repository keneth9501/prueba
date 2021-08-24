# import requests module
import requests
import pandas as  pd
import os
import json
import pandas as pd
from pandas.io.json import json_normalize


# Making a get request
#response = requests.get('https://multipagos-team.myfreshworks.com/crm/sales/api/selector/contact_statuses',
#headers={'Authorization': 'Token token=g64t0eqabGOEijISr2lliw'})
  
# print response
#print(response)
  
# print json content


#response = requests.get('https://multipagos-team.myfreshworks.com/crm/sales/api/selector/currencies',
#headers={'Authorization': 'Token token=g64t0eqabGOEijISr2lliw'})

#print(response.json())

response = requests.get('https://multipagos-team.myfreshworks.com/crm/sales/api/selector/owners',
headers={'Authorization': 'Token token=g64t0eqabGOEijISr2lliw'})

print(response.json())

response = json.loads(requests.get('https://multipagos-team.myfreshworks.com/crm/sales/api/contacts/view/16001797531?include=owner',
headers={'Authorization': 'Token token=g64t0eqabGOEijISr2lliw'}).text)

df_nested_list = pd.json_normalize(response, record_path = ['contacts'])

df_nested_list.to_excel(r'C:\Users\50576\Documents\00002021.xlsx',engine='xlsxwriter',
index=False,sheet_name='pagos')

print(df_nested_list)
