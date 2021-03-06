import os
import json
import sys
import requests
from pprint import pprint

from util import read_sysargv

chamber, congress = read_sysargv()

def import_json(chamber, congress):
    pathname = f'./data/{chamber}_{congress}.json'
    if os.path.isfile(pathname):
        print("File already exists")
        sys.exit()

    host = "https://api.propublica.org/congress/v1/"
    url = host+"{}/{}/members.json".format(congress, chamber)
    api_key = os.environ.get('CONGRESS_API_KEY')

    if not api_key:
        raise Exception("No key present! Run 'source env.sh' to load")

    headers = { 'X-API-Key': api_key }
    r = requests.get(url, headers=headers)
    r.raise_for_status()

    data = r.json()

    with open(pathname, 'w') as outfile:  
        json.dump(data, outfile)

if chamber in ['both', 'house']:
    import_json('house', congress)
if chamber in ['both', 'senate']:
    import_json('senate', congress)

print("Import complete")
