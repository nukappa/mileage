# Mileage calculator

## About
This `Python` script can be used to compute the driving mileages of each
month. It uses the [https://github.com/googlemaps/google-maps-services-python](googlemaps)
library, which in turn is based on the Google Maps APIs.

## Installation, prerequisites and usage
Just clone or copy the script locally and run it (Python 2.7+ required).
You will need to get an API key in order to use the Google Maps APIs, see 
[https://github.com/googlemaps/google-maps-services-python](here) for more
details. The places are stored as dictionaries, for example
```
work = {'name' : 'Work', 'address' : 'Platz der Republik 1, 11011'}
hbf = {'name' : 'Hauptbahnhof', 'address' : 'Europaplatz, 10557'}
```
and there is a function which calculates and returns the distances between
the places as a route
```
add_entry([work, hbf, work])
```
After creating the list of routes, the whole information is then saved as an 
`xlsx` sheet by using the `xlsxwriter` library. The script calculates the total
number of kms, subtracts the private drives and introduces dates where appropriate.
See the example included for more details.
