import requests
import os
from pprint import pprint

host = "https://api.propublica.org/congress/v1/"
congress = "116"
chamber = "senate"
url = host+"{}/{}/members.json".format(congress, chamber)
api_key = os.environ.get('CONGRESS_API_KEY')

if not api_key:
    raise Exception('No key present! Run source env.sh to load')

headers = { 'X-API-Key': api_key }
r = requests.get(url, headers=headers)
r.raise_for_status()

data = r.json()

index = 0
for item in data['results']['members']:
    index += 1
    # print(item)
    if index % 10 == 0:
        # print('st')
        pprint(item)
    item.get('first_name')
    item.get('last_name')
print(data.keys())
# pprint(data)