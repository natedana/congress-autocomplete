import sys
import os
import json

def to_num(val):
    try:
        return int(val)
    except (ValueError, TypeError):
        return val.lower()

def read_sysargv():
    args = list(sys.argv)

    default_chamber = 'senate'
    default_congress = 116
    congress = None
    chamber = None
    force = False
        
    if len(args) > 1:
        if '-f' in args:
            force = True 
            args.remove('-f')

        if len(args) == 2:
            arg1 = to_num(args[1])
            if isinstance(arg1, str):
                chamber = arg1
                congress = default_congress
            else:
                chamber = default_chamber
                congress = arg1
        elif len(args) == 3:
            arg1 = to_num(args[1])
            arg2 = to_num(args[2])
            if isinstance(arg1, str) and isinstance(arg2, int):
                chamber = arg1
                congress = arg2
            else:
                chamber = arg2
                congress = arg1
        else:
            raise Exception('Too many args')
    else:
        chamber = default_chamber
        congress = default_congress

    if chamber.lower() not in ['senate', 'house', 'both']:
        raise Exception('Invalid chamber')
    if congress > 116 or congress < 100:
        raise Exception('Invalid congress')

    return (chamber, congress, force)

def load_data_file(chamber, congress):
    if chamber not in ['senate', 'house', 'both']:
        raise Exception('Need valid chamber value, got: ', chamber)

    house = False
    senate = False

    data = {}

    if chamber == 'both':
        house = True
        senate = True
    elif chamber == 'senate':
        senate = True
    elif chamber == 'house':
        house = True

    if house:
        house_filename = f'./data/house_{congress}.json'
        house_data = os.path.isfile(house_filename)
        if not house_data:
            print("No house data present for {} congress.".format(congress))

        with open(house_filename, encoding='utf-8') as json_file:  
            house_data = json.loads(json_file.read())
            house_data = house_data['results'][0]['members']
            data['house'] = house_data

    if senate:
        senate_filename = f'./data/senate_{congress}.json'
        senate_data = os.path.isfile(senate_filename)

        if not senate_data:
            print("No senate data present for {} congress.".format(congress))

        with open(senate_filename, encoding='utf-8') as json_file:  
            senate_data = json.loads(json_file.read())
            senate_data = senate_data['results'][0]['members']
            data['senate'] = senate_data

    return data


def parse_senator(senator):
    first_name = senator['first_name']
    last_name = senator['last_name']
    party = senator['party']
    state = senator['state']

    snippet = f'{last_name}, {first_name}'
    full_title = f'Sen. {first_name} {last_name} ({party}-{state})'

    return (snippet, full_title)

def parse_house_member(member):
    first_name = member['first_name']
    last_name = member['last_name']
    short_title = member['short_title']
    district = member['district']
    state = member['state']
    party = member['party']

    snippet = f'{last_name}, {first_name}'
    full_title = f'{short_title} {first_name} {last_name} ({party}-{state}-{district})'

    return (snippet, full_title)