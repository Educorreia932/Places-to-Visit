# -*- coding: utf-8 -*-
"""
Created on Sun May 26 21:31:44 2019

@author: Utilizador
"""

import time

start = time.time()

from geopy.geocoders import Nominatim

geolocator = Nominatim(user_agent ="my-application")

import xml.etree.ElementTree as ET

root = ET.parse('places.kml').getroot()

Document = root[0]

places = []

for label in Document:
    if label.tag[32:] == "Folder":
        for place in label:    
            if place.tag[32:] == "Placemark":
                for attribute in place:
                    if attribute.tag[32:] == "name":
                        name = attribute.text
                        
                    elif attribute.tag[32:] == "styleUrl":
                        if attribute.text == "#icon-1899-E65100-nodesc":
                            already_visited = "Não"
                            
                        elif attribute.text == "#icon-1899-0F9D58-nodesc":
                            already_visited = "Sim"
                            
                        else:
                            already_visited = "N/A"
                        
                    elif attribute.tag[32:] == "Point":
                        coordinates = eval(attribute[0].text.strip()[:-2])[::-1]                        
                        
                        place = geolocator.reverse(coordinates, timeout = 15, language = "pt-pt").raw     
                        country = place["address"]["country"]
                        
                        if "suburb" in place["address"].keys():
                            locality = place["address"]["suburb"]
                            
                        elif "city_district" in place["address"].keys():
                            locality = place["address"]["city_district"]
                            
                        else:
                            locality = "N/A"
                    
                places.append([country, locality, name, None, already_visited, None])
                        
import xlsxwriter

# Create an new Excel file and add a worksheet.
workbook = xlsxwriter.Workbook('places.xlsx')
worksheet = workbook.add_worksheet()

# Widen the columns to make the text clearer.
worksheet.set_column("A:A", 23)
worksheet.set_column("B:B", 20)
worksheet.set_column("C:C", 53)
worksheet.set_column("E:E", 11)
worksheet.set_column("F:F", 13)

# Add a bold format to use to highlight cells.
bold = workbook.add_format({'bold': True})

worksheet.write(0, 0, "País", bold)
worksheet.write(0, 1, "Localidade", bold)
worksheet.write(0, 2, "Nome", bold)
worksheet.write(0, 3, "Tipo", bold)
worksheet.write(0, 4, "Já visitado", bold)
worksheet.write(0, 5, "Observações", bold)

counter = 1

for place in places:
    worksheet.write(counter, 0, place[0])
    worksheet.write(counter, 1, place[1])
    worksheet.write(counter, 2, place[2])
    worksheet.write(counter, 4, place[4])
    
    counter += 1

workbook.close()

end = time.time() 

print("Finished with success, it took " + str(end - start) + " seconds to finish.")
    
