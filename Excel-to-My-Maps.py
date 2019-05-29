    # -*- coding: utf-8 -*-
"""
Created on Tue May 28 21:34:04 2019

@author: Utilizador
"""

import time
from geopy.geocoders import Nominatim
from googletrans import Translator
import xlrd 

start = time.time()
  
translator = Translator()
geolocator = Nominatim(user_agent ="Excel-to-My-Maps")

file_name = "places.xlsx"

# Give the location of the file 
loc = (file_name) 
  
workbook = xlrd.open_workbook(loc) 
sheet = workbook.sheet_by_index(0) 
sheet.cell_value(0, 0) 
  
counter = 0

for i in range(1, sheet.nrows):
    name = sheet.cell_value(i, 2)
    
    location = geolocator.geocode(name, timeout = 15, addressdetails = True)
    
    if location == None:
        print(name)
        counter += 1


#places.append([country, locality, name, category, already_visited, None])
places = []

print(counter)

end = time.time() 

print("Finished with success, it took " + str(end - start) + " seconds to finish.")