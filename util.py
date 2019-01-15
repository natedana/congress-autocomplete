import sys

def to_num(val):
    num = None
    try:
        num = int(val)
    except ValueError:
        try:
            num = float(val)
        except ValueError:
            num = val
    return num.lower()

def read_sysargv():

    default_chamber = 'senate'
    default_congress = 116
    congress = None
    chamber = None
        
    if len(sys.argv) > 1:
        if len(sys.argv) == 2:
            arg1 = to_num(sys.argv[1])
            if isinstance(arg1, str):
                chamber = arg1
                congress = default_congress
            else:
                chamber = default_chamber
                congress = arg1
        elif len(sys.argv) == 3:
            arg1 = to_num(sys.argv[1])
            arg2 = to_num(sys.argv[2])
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

    if chamber.lower() not in ['senate', 'house']:
        raise Exception('Invalid chamber')
    if congress > 116 or congress < 100:
        raise Exception('Invalid congress')

    return (congress, chamber)

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