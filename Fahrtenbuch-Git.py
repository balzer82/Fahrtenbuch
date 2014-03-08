# -*- coding: utf-8 -*-
# <nbformat>3.0</nbformat>

# <codecell>

import requests as r
import json
import time
import xlsxwriter

# <headingcell level=1>

# Fahrtenbuch aus MOVES App

# <codecell>

detail = 1 # the lower, the more precise the location is geoencoded
fromdate = '20140201' # date in yyyyMMdd
todate = '20140228' # date in yyyyMMdd

# <codecell>

# Source: https://gist.github.com/bradmontgomery/5397472
def getadress(latitude,longitude):
    # grab some lat/long coords from wherever. For this example,
    # I just opened a javascript console in the browser and ran:
    #
    # navigator.geolocation.getCurrentPosition(function(p) {
    #   console.log(p);
    # })
    #
    #latitude = 35.1330343
    #longitude = -90.0625056
 
    # Did the geocoding request comes from a device with a
    # location sensor? Must be either true or false.
    sensor = 'true'
 
    # Hit Google's reverse geocoder directly
    # NOTE: I *think* their terms state that you're supposed to
    # use google maps if you use their api for anything.
    base = "http://maps.googleapis.com/maps/api/geocode/json?"
    params = "latlng={lat},{lon}&sensor={sen}".format(
        lat=latitude,
        lon=longitude,
        sen=sensor
    )
    url = "{base}{params}".format(base=base, params=params)

    response = r.get(url)
    
    if response.status_code==200:
        pass
    else:
        print('Error: \'%s\'' % response.json()['error'])
    
    return response.json()['results'][detail]['formatted_address']

# <codecell>

url = 'https://api.moves-app.com/api/1.1' # Moves API
params=  {'access_token': '********************', # your Access Token here
          'trackPoints': 'false',
          'from': fromdate,
          'to': todate}
            #'pastDays': '10',
timeline = r.get(url+'/user/storyline/daily', params=params)

if timeline.status_code==200:
    print('Received Data from Moves API.')
else:
    print('Error: \'%s\'' % timeline.json()['error'])

# <codecell>

begin = time.strptime(timeline.json()[0]['date'], '%Y%m%d')
ende = time.strptime(timeline.json()[-1]['date'], '%Y%m%d')
print('Zeitraum: %s - %s' % (time.strftime("%A, %d %b %Y", begin), time.strftime("%A, %d %b %Y", ende)))

# <headingcell level=2>

# Create Excel File

# <codecell>

# Create an new Excel file and add a worksheet.
workbook = xlsxwriter.Workbook('Fahrtenbuch-%s.xlsx' % time.strftime("%Y", begin))
worksheet = workbook.add_worksheet(time.strftime("%b-%Y", begin))
# Add a bold format to use to highlight cells.
bold = workbook.add_format({'bold': True})

# Text with formatting.
worksheet.write(0,0, 'Fahrtenbuch', bold)
worksheet.write(2,0, 'Von:')
worksheet.write(2,1, '%s' % time.strftime("%A, %d %b %Y", begin))
worksheet.write(3,0, 'Bis:')
worksheet.write(3,1, '%s' % time.strftime("%A, %d %b %Y", ende))

# <headingcell level=2>

# Crawl the Data

# <codecell>

fahrtenbuchdata=[]

# <codecell>

for day in range(len(timeline.json())): # Tage durchgehen
    
    # Datum
    date = time.strptime(timeline.json()[day]['date'], '%Y%m%d')
    print('%s' % (time.strftime("%A, %d %b %Y", date)))
    print(20*'=')
    
    fahrtenbuchdata.append('%s' % (time.strftime("%A, %d %b %Y", date)))
    
    transport=False
    for i in range(len(timeline.json()[day]['segments'])): # Segmente durchgehen
        
        try: # Ort
            place = timeline.json()[day]['segments'][i]['place']
            lat = timeline.json()[day]['segments'][i]['place']['location']['lat']
            lon = timeline.json()[day]['segments'][i]['place']['location']['lon']
            location= getadress(lat, lon)
            if timeline.json()[day]['segments'][i]['place'].has_key('name'):
                name = timeline.json()[day]['segments'][i]['place']['name']
                
                if transport:
                    #print('%s (%s)\n' % (location, name))
                    print('%s\n' % (location))
                    fahrtenbuchdata.append('%s' % (location))
                    fahrtenbuchdata.append('')
                    transport=False
            else:
                if transport:
                    print('%s\n' % (location))
                    fahrtenbuchdata.append('%s' % (location))
                    fahrtenbuchdata.append('')
                    transport=False
                pass
        except: # an sonsten ist es eine Verbindung
            try:
                #print '.'
                for act in range(len(timeline.json()[day]['segments'][i]['activities'])):

                    # nur Transport nutzen
                    if timeline.json()[day]['segments'][i]['activities'][act]['activity']=='transport':
                        starttime = time.strptime(timeline.json()[day]['segments'][i]['activities'][act]['startTime'], "%Y%m%dT%H%M%S+%f")
                        distance= float(timeline.json()[day]['segments'][i]['activities'][act]['distance'])
                        endtime = time.strptime(timeline.json()[day]['segments'][i]['activities'][act]['endTime'], "%Y%m%dT%H%M%S+%f")
                        
                        #print('%s (%s)' % (location, name))
                        print('%s' % (location))
                        fahrtenbuchdata.append('%s' % (location))
                        route = '%.1fkm (%s - %s)' % (distance/1000.0, time.strftime("%H:%MUhr", starttime), time.strftime("%H:%MUhr", endtime))
                        print(route)
                        fahrtenbuchdata.append(route)
                        
                        transport=True
                    else:
                        pass
                #print '.'
            except: # oder was anderes
                pass

    print('\n')
    fahrtenbuchdata.append('')

# <headingcell level=2>

# Save to Excel File

# <codecell>

for i in range(len(fahrtenbuchdata)):
    worksheet.write(i+5, 0, fahrtenbuchdata[i])

# <codecell>

print('Erledigt. Nun beten und hoffen, dass das Finanzamt es akzeptiert.')

