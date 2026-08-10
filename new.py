import requests
import urllib3
urllib3.disable_warnings()

session = requests.Session()
session.auth = ("lop2cob", "shreyansh4991Ab#") # Your credentials
headers = {"Accept": "application/rdf+xml", "OSLC-Core-Version": "2.0"}

# Sureshkumar's URL
response = session.get("https://rb-alm-06-p.de.bosch.com/jts/users/QRB1COB", headers=headers, verify=False)

print(response.text)
