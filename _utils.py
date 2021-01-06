import googlemaps
from math import ceil
import numpy as np

def read_gmaps_key(file):
    '''Reads a googlemaps key from the disk and returns a client'''
    with open(file, 'r') as f:
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
        return 'No ride', 0
    else:
        distance = np.random.randint(5, 24)
        return 'private', distance

def add_entry(route, distance='', description=''):
    """Returns a row for the table."""
    if route == 'private' or route == 'no':
        return (route + ' drive', distance, '')
    distance = get_route_distance(route)
    route = [place['name'] for place in route]
    route = ' -- '.join(route)

    return route, distance, description