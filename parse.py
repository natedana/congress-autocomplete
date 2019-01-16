import json
import os
import sys

from pprint import pprint

from util import read_sysargv, parse_senator, parse_house_member, load_data_file
from file_generator import File

chamber, congress = read_sysargv()
data = load_data_file(chamber, congress)

filename = f'autotext_{chamber}_{congress}.bas'
macro = File(filename)

def process(chamber):
    for _, member in enumerate(data.get(chamber, [])):
        parse = parse_senator if chamber == 'senate' else parse_house_member
        snippet, title = parse(member)

        macro.add_autotext(snippet, title)

process('house')
process('senate')

macro.close_file()

print("File generation complete: saved file as {}".format(filename))
    