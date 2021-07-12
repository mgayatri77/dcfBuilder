import urllib.request, json 
with urllib.request.urlopen("https://data.sec.gov/api/xbrl/companyconcept/CIK0001341439/us-gaap/AccountsPayableCurrent.json") as url:
    oracle_data = json.loads(url.read().decode())  
    print(oracle_data)
