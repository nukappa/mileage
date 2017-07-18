#!/usr/bin/env python

__author__ = "Nikos Karaiskos"
__copyright__ = "MIT License, Copyright (c) 2007"
__version__ = "0.1.0"

import googlemaps
import xlsxwriter
from math import ceil

# The google API key
gmaps = googlemaps.Client(key='insert_your_key_here')

def get_distance(address_1, address_2, mode='driving'):
    """Computes the distance between two addresses for the requested way of
    travelling. The distance is returned in km, rounded to the next integer.

    Arguments:
    mode -- can be 'driving', 'transit', 'walking', 'bicycling'
    """
    route = gmaps.directions(address_1, address_2, mode=mode)
    distance = route[0]['legs'][0]['distance']['text']
    distance = str(distance).split(' ')[0]

    if ',' in distance:
        return int(''.join(distance.split(',')))
    else:
        return int(ceil(float(distance)))

def get_route_distance(route):
    """Computes the total distance for a route of multiple destinations. The
    route is a list of places, each place being a dictionary with a name and
    an address.
    """
    distance = 0
    for idx in range(len(route)-1):
        distance += get_distance(route[idx]['address'], route[idx+1]['address'])
 
    return distance

def add_entry(route, distance='', description=''):
    """Returns a row for the table."""
    if route == 'private' or route == 'no':
        return (route + ' drive', distance, '')
    distance = get_route_distance(route)
    route = [place['name'] for place in route]
    route = ' -- '.join(route)

    return (route, distance, description)


# The months
months = ['January', 'February', 'March', 'April', 'May', 'June', 'July',
          'August', 'September', 'October', 'November', 'December']

# Here define more places
home = {'name' : 'Home', 'address' : 'Bodestrasse 1-3, 10178'}
work = {'name' : 'Work', 'address' : 'Platz der Republik 1, 11011'}
hbf = {'name' : 'Hauptbahnhof', 'address' : 'Europaplatz, 10557'}

# Enter basic information
last_km_stand = 89768
current_month = 'July'
current_year = '2017'
month_idx = months.index(current_month)+1
previous_month = months[months.index(current_month)-1]

# Enter the routes as a list
the_list = [add_entry('private', 12),
            add_entry([home, work, home]),
            add_entry([home, work, hbf, work, home], description='Picking up client'),
            add_entry('no', 0),
            add_entry('private', 2),
            add_entry('private', 16)]

# Calculate the total numbers of kms
private_kms = 0
normal_kms = 0
for entry in the_list:
    if entry[0] == 'private drive':
        private_kms += entry[1]
    else:
        normal_kms += entry[1]

# Create the excel sheet
workbook = xlsxwriter.Workbook('mileage_' + current_year + '_' + 
                                current_month +  '.xlsx')
worksheet = workbook.add_worksheet()
worksheet.set_column(0, 0, 15)
worksheet.set_column(1, 1, 50)
worksheet.set_column(2, 2, 10)
worksheet.set_column(3, 3, 40)

# Format styles
bold = workbook.add_format({'bold': True})
bold_center = workbook.add_format({'bold': True, 'align' : 'center'})
bold_right = workbook.add_format({'bold': True, 'align' : 'right'})
center = workbook.add_format({'align' : 'center'})
red_text = workbook.add_format({'font_color' : 'red'})
red_text_center = workbook.add_format({'font_color' : 'red', 'align' : 'center'})
red_text_right = workbook.add_format({'font_color' : 'red', 'align' : 'right', 'bold' : 'True'})

# Write the header
worksheet.write('A1', current_month + ' ' + current_year, bold_center)
worksheet.write('B1', 'Last mileage from ' + previous_month, bold)
worksheet.write('C1', last_km_stand, bold_center)
worksheet.write('A3', 'Date', bold_center)
worksheet.write('B3', 'Route', bold_center)
worksheet.write('C3', 'Km', bold_center)
worksheet.write('D3', 'Comments', bold)

# Write the entries of the month
for entry in range(len(the_list)):
    worksheet.write('A' + str(4+entry), str(entry+1)+'/' + str(month_idx) 
                        + '/' + current_year, center)
    if the_list[entry][0] == 'private drive':
        worksheet.write('B' + str(4+entry), the_list[entry][0], red_text)
        worksheet.write('C' + str(4+entry), the_list[entry][1], red_text_center)
    else:
        worksheet.write('B' + str(4+entry), the_list[entry][0])
        worksheet.write('C' + str(4+entry), the_list[entry][1], center)
    worksheet.write('D' + str(4+entry), the_list[entry][2])

# Write the footer
worksheet.write('B' + str(4+len(the_list)+1), 'Overall km:', bold_right)
worksheet.write('C' + str(4+len(the_list)+1), normal_kms+private_kms, center)
worksheet.write('B' + str(4+len(the_list)+2), 'of which private:', red_text_right)
worksheet.write('C' + str(4+len(the_list)+2), private_kms, red_text_center)
worksheet.write('B' + str(4+len(the_list)+3), 'Tax deductible:', bold_right)
worksheet.write('C' + str(4+len(the_list)+3), normal_kms, center)

workbook.close()