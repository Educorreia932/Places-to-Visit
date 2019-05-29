# -*- coding: utf-8 -*-
"""
Created on Sun May 26 21:31:44 2019

@author: Utilizador
"""

import time
from geopy.geocoders import Nominatim
import xml.etree.ElementTree as ET
import xlsxwriter

start = time.time()

geolocator = Nominatim(user_agent ="My-Maps-to-Excel")

file_name = "places.kml"

root = ET.parse(file_name).getroot()

Document = root[0]

categories = {"theme_park": "Parque temático",
              "cafe": "Café",
              "restaurant": "Restaurante",
              "zoo": "Zoo",
              "attraction": "Atração",
              "building": "Edifício",
              "information": "Informação",
              "mall": "Centro comercial",
              "bar": "Bar"}

places = []

for label in Document:
    if label.tag[32:] == "Folder":
        for place in label:    
            if place.tag[32:] == "Placemark":
                for attribute in place:
                    #Name of place
                    if attribute.tag[32:] == "name":
                        name = attribute.text
                    
                    #Check if already visited or not
                    elif attribute.tag[32:] == "styleUrl":
                        if attribute.text == "#icon-1899-E65100-nodesc":
                            already_visited = "Não"
                            
                        elif attribute.text == "#icon-1899-0F9D58-nodesc":
                            already_visited = "Sim"
                            
                        else:
                            already_visited = "N/A"
                        
                    #Address details
                    elif attribute.tag[32:] == "Point":
                        coordinates = eval(attribute[0].text.strip()[:-2])[::-1]                        
                        place = geolocator.reverse(coordinates, addressdetails = True, timeout = 15, language = "pt-pt").raw 
                        
                        country = place["address"]["country"]    
                                     
                        
                        if "suburb" in place["address"]:
                            locality = place["address"]["suburb"]
                            
                        elif "city_district" in place["address"]:
                            locality = place["address"]["city_district"]
                            
                        elif "state" in place["address"]:
                            locality = place["address"]["state"]
                            
                        else:
                            locality = "N/A"
                        
                        category = "N/A"
                            
                        location = geolocator.geocode(name, timeout = 15, addressdetails = True)
                        
                        if location is not None:
                            category = location.raw["type"]

                places.append([country, locality, name, category, already_visited, None])
                        
# Create an new Excel file and add a worksheet.
workbook = xlsxwriter.Workbook('places.xlsx')
worksheet = workbook.add_worksheet()

# Widen the columns to make the text clearer.
worksheet.set_column("A:A", 23)
worksheet.set_column("B:B", 20)
worksheet.set_column("C:C", 53)
worksheet.set_column("D:D", 19)
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
    worksheet.write(counter, 3, place[3])
    worksheet.write(counter, 4, place[4])
    
    counter += 1

workbook.close()

end = time.time() 

print("Finished with success, it took " + str(end - start) + " seconds to finish.")
    
