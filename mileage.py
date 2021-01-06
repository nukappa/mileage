__author__ = "Nikos Karaiskos"
__copyright__ = "MIT License, Copyright (c) 2007"
__version__ = "0.2"

from _utils import *

month_data_file = 'data/example.yaml'

gmaps = read_gmaps_key('data/gmaps_key.txt')
months = read_months('data/months_en.txt')
month_data = read_month_data(month_data_file)
month_data = add_month_metadata(months, month_data)
db = read_database('data/addresses.json')

route_list = create_route_list(gmaps, db, month_data)
write_sheet(route_list, month_data)

write_month_data(route_list, month_data, month_data_file)