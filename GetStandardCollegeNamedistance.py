import json
import ast
import pygeohash as pgh
import requests
import urllib
import geopy
from geopy import distance

base_url = "https://maps.googleapis.com/maps/api/geocode/json?"
AUTH_KEY = "AIzaSyDKLCOQ7zgXOQZqIvp853W6dXU1XMktmtk"

def GetStandardName(data):
    parameters = {
        "address": data,
        "key": AUTH_KEY
    }
    if data in ['0',0,None,"#N/A",'undefined']:
        return data
    with open('address_college.txt','r') as f:
        file = f.read()
        fields = ast.literal_eval(file)
        address_college = fields
    with open('college_distance.txt','r') as f:
        file = f.read()
    #
    for colls in address_college:
        colleges = address_college[colls]
        if data in colleges:
            address_college[colls].append(data)
            with open('address_college.txt', 'w') as f:
                f.write(json.dumps(address_college))
            return colleges[0]

    r = json.loads(requests.get(f"{base_url}{urllib.parse.urlencode(parameters)}").text)
    try:
        standard_address = r["results"][0]["address_components"][6]["long_name"]
        if standard_address not in address_college:
            address_college[standard_address] = [data]
        else:
            address_college[standard_address].append(data)
        with open('address_college.txt', 'w') as f:
            f.write(json.dumps(address_college))

        return address_college[standard_address][0]
    except:
        pass
    return
def GetDistance(data):
    parameters = {
        "address": data,
        "key": AUTH_KEY,
    }
    if data in ['0',0,None,"#N/A",'undefined']:
        return

    college_distance = {}
    with open('college_distance.txt','r') as f:
        file = f.read()
        college_distance = ast.literal_eval(file)
    if data in ['0',0,None,"#N/A",'undefined']:
        return None
    for college in college_distance:
        if data == college:
            if int(college_distance[data]) > 3000:
                return
            return college_distance[data]

    r = json.loads(requests.get(f"{base_url}{urllib.parse.urlencode(parameters)}").text)
    try:
        lat1 = 13.010086
        long1 = 77.55106
        chord1 = (lat1, long1)

        # Co - ordinates of place given it is not a garbage value
        lat2 = r['results'][0]['geometry']['location']['lat']
        long2 = r['results'][0]['geometry']['location']['lng']
        chord2 = (lat2, long2)
        dist = distance.distance(chord1, chord2).km
        formatted_address = r['results'][0]['formatted_address']

        if data not in college_distance:
            college_distance[data] = int(dist)
        with open('college_distance.txt', 'w') as f:
            f.write(json.dumps(college_distance))
        if int(dist) > 3000:
            return
        return int(dist)
    except:
        pass
    return

def GetParticipants(college):
    with open('address_college.txt','r') as f:
        file = f.read()
        fields = ast.literal_eval(file)
        address_college = fields
    for address in address_college:
        if college in address_college[address]:
            return len(address_college[address])
    return 1
# dist = GetDistance('JBU')
# print(dist)