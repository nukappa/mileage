from calendar import monthrange
import datetime
import googlemaps
import json
from math import ceil
import numpy as np
import subprocess
import xlsxwriter
import yaml

def read_months(filename):
    '''Read list of months from disk and return a list'''
    with open(filename, 'r') as f:
        months = [line.strip() for line in f]
    return months

def read_database(filename):
    '''Read local database of addresses from disk and return it as dictionary'''
    db = json.load(open(filename, 'r'))
    return db

def read_month_data(filename):
    '''Read the month data from disk and return it as dictionary'''
    with open(filename, 'r') as stream:
        try:
            month_data = yaml.safe_load(stream)
        except yaml.YAMLError as exc:
            print(exc)
    return month_data

def write_month_data(route_list, month_data, filename):
    all_kms = sum(calculate_kms(route_list))
    month_data['new_km_stand'] = month_data['last_km_stand'] + all_kms
    with open(filename, 'w') as f:
        yaml.dump(month_data, f)

def add_month_metadata(months, month_data):
    month_data['month_idx'] = months.index(month_data['month'])+1
    month_data['prev_month'] = months[months.index(month_data['month'])-1]
    month_data['days'] = monthrange(month_data['year'], month_data['month_idx'])[1]
    return month_data


def read_gmaps_key(filename):
    '''Reads a googlemaps key from the disk and returns a client'''
    with open(filename, 'r') as f:
        gmaps_key = f.readlines()[0]
    gmaps = googlemaps.Client(key=gmaps_key)  
    return gmaps

def get_distance(gmaps, address_1, address_2, mode='driving'):
    '''Computes the distance between two addresses for the requested way of
    travelling. The distance is returned in km, rounded to the next integer.

    Args:
        gmaps: a googlemaps client
        address_1: string
        address_2: string
        mode: string, can be 'driving' (default), 'transit', 'walking', 'bicycling'

    Returns:
        int, the distance in km.
    '''
    route = gmaps.directions(address_1, address_2, mode=mode)
    distance = route[0]['legs'][0]['distance']['text']
    distance = str(distance).split(' ')[0]

    if ',' in distance:
        return int(''.join(distance.split(',')))
    else:
        return int(ceil(float(distance)))

def get_route_distance(gmaps, route):
    """Computes the total distance for a route of multiple destinations. The
    route is a list of places, each place being a dictionary with a name and
    an address.

    Args:
        gmaps: a googlemaps client
        route: list of strings (addresses)

    Returns:
        distance: int, total distance in km.
    """
    distance = 0
    for idx in range(len(route)-1):
        distance += get_distance(gmaps, route[idx]['address'],
                                 route[idx+1]['address'])
 
    return distance

def weekend_ride():
    choice = np.random.randint(2)
    if choice == 0:
        return 'Keine', 0
    else:
        distance = np.random.randint(5, 24)
        return 'Private', distance

def add_entry(gmaps, route, distance='', description=''):
    """Returns a row for the table."""
    if route[0] == 'Keine':
        return (route[0] + ' Fahrt', 0, '')
    if route[0] == 'Private':
        return (route[0] + ' Fahrt', route[1], '')
    distance = get_route_distance(gmaps, route)
    route = [place['name'] for place in route]
    route = ' - '.join(route)

    return route, distance, description


def create_route_list(gmaps, db, month_data):
    '''Create the list of routes by taking into account exceptions.

    Args:
        gmaps: a googlemaps client
        month_data: dictionary holding the data of the month

    Returns:
        route_list: a list of routes
    '''
    route_list = []
    for day in range(1, month_data['days']+1):
        if day in month_data['exceptions']:
            route = month_data['exceptions'][day]
            if route[0] == 'Keine':
                route_list.append(add_entry(gmaps, route))
            elif route[0] == 'Private':
                route_list.append(add_entry(gmaps, route))
            else:
                route_list.append(add_entry(gmaps, [db[address] for address in route]))
        else:
            date = datetime.datetime(month_data['year'], month_data['month_idx'], day)
            if date.weekday() >= 5:
                route_list.append(add_entry(gmaps, weekend_ride()))
            else:
                route_list.append(add_entry(gmaps, [db['Haus'], db['Praxis'], db['Haus']]))
    return route_list

def calculate_kms(route_list):
    '''Calculate and return the normal and private kms'''
    private_kms = 0
    normal_kms = 0
    for entry in route_list:
        if entry[0] == 'Private Fahrt':
            private_kms += entry[1]
        else:
            normal_kms += entry[1]
    return normal_kms, private_kms

def write_sheet(route_list, month_data):
    '''Create the excel sheet'''
    normal_kms, private_kms = calculate_kms(route_list)

    # Create the excel sheet
    workbook = xlsxwriter.Workbook('out/Fahrbuch_' + str(month_data['year']) + '_'
        + month_data['month'] +  '.xlsx')
    worksheet = workbook.add_worksheet()
    worksheet.set_landscape()
    worksheet.set_margins(top=0.35, bottom=0.35, right=0.5, left=0.35)
    worksheet.set_column(0, 0, 15)
    worksheet.set_column(1, 1, 40)
    worksheet.set_column(2, 2, 10)
    worksheet.set_column(3, 3, 20)

    # Format styles
    bold = workbook.add_format({'bold': True})
    bold_left = workbook.add_format({'bold': True, 'align' : 'left'})
    bold_center = workbook.add_format({'bold': True, 'align' : 'center'})
    bold_right = workbook.add_format({'bold': True, 'align' : 'right'})
    center = workbook.add_format({'align' : 'center'})
    red_text = workbook.add_format({'font_color' : 'red'})
    red_text_center = workbook.add_format({'font_color' : 'red', 'align' : 'center'})
    red_text_right = workbook.add_format({'font_color' : 'red', 'align' : 'right', 'bold' : 'True'})

    # Write the header
    worksheet.write('A1', month_data['month'] + ' ' + str(month_data['year']), bold_center)
    worksheet.write('B1', 'Last mileage from ' + month_data['prev_month'], bold)
    worksheet.write('C1', month_data['last_km_stand'], bold_center)
    worksheet.write('A2', 'Date', bold_center)
    worksheet.write('B2', 'Route', bold_left)
    worksheet.write('C2', 'Km', bold_center)
    worksheet.write('D2', 'Comments', bold)

    # Write the entries of the month
    for entry in range(len(route_list)):
        worksheet.write('A' + str(3+entry), str(entry+1)+'/' 
                        + str(month_data['month_idx']) 
                        + '/' + str(month_data['year']), center)
        if route_list[entry][0] == 'Private Fahrt':
            worksheet.write('B' + str(3+entry), route_list[entry][0], red_text)
            worksheet.write('C' + str(3+entry), route_list[entry][1], red_text_center)
        else:
            worksheet.write('B' + str(3+entry), route_list[entry][0])
            worksheet.write('C' + str(3+entry), route_list[entry][1], center)
        worksheet.write('D' + str(3+entry), route_list[entry][2])

    # Write the footer
    worksheet.write('B' + str(3+len(route_list)), 'Overall km:', bold_right)
    worksheet.write('C' + str(3+len(route_list)), normal_kms + private_kms, center)
    worksheet.write('B' + str(3+len(route_list)+1), 'of which private:', red_text_right)
    worksheet.write('C' + str(3+len(route_list)+1), private_kms, red_text_center)
    worksheet.write('B' + str(3+len(route_list)+2), 'Tax deductible:', bold_right)
    worksheet.write('C' + str(3+len(route_list)+2), normal_kms, center)

    workbook.close()

    subprocess.call(['libreoffice', '--headless',
        '--convert-to', 'pdf',
        'out/Fahrbuch_' + str(month_data['year']) + '_'
        + month_data['month'] +  '.xlsx',
        '--outdir', './out/'])