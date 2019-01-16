import json
import os
import sys

from pprint import pprint

from util import read_sysargv, parse_senator, parse_house_member, load_data_file
from file_generator import BasFile

chamber, congress, force = read_sysargv()
data = load_data_file(chamber, congress)

filename = f'autotext_{chamber}_{congress}.bas'
macro = BasFile(filename, force)

def process(chamber):
    for _, member in enumerate(data.get(chamber, [])):
        parse = parse_senator if chamber == 'senate' else parse_house_member
        snippet, title = parse(member)

        macro.add_autotext(snippet, title)

process('house')
process('senate')

macro.close_file()

print("File generation complete: saved file as {}".format(filename))
    