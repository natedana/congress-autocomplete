import json
import os
import sys
from pprint import pprint
from util import read_sysargv, parse_senator, parse_house_member

congress, chamber = read_sysargv()

pathname = f'./data/{chamber}_{congress}.json'

if not os.path.isfile(pathname):
    print(f'File: {pathname} does not exist')
    sys.exit()

with open(f'./data/{chamber}_{congress}.json', encoding='utf-8') as json_file:  
    data = json.loads(json_file.read())

members = data['results'][0]['members']

titles = []

for index, member in enumerate(members):

    parse = parse_senator if chamber == 'senate' else parse_house_member
    titles.append(parse(member))
    snippet, title = parse(member)

    code = "Set oAutoText = Templates(ActiveDocument.AttachedTemplate).AutoTextEntries.Add(Name:='{}', Value:= '{}', Range:=Selection.Range)".format(snippet, title)

print()
    