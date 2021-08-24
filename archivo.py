import requests
import json

url = "https://multipagos-team.myfreshworks.com/crm/sales/api/selector/owners"

payload="{\"query\":\"\",\"variables\":{}}"
headers = {
  'Authorization': 'Token token=g64t0eqabGOEijISr2lliw',
  'Content-Type': 'application/json'
}

response = requests.request("GET", url, headers=headers, data=payload)

print(response.text)

