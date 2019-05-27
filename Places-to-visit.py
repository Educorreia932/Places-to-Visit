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

#Correspondency between country name in original language and PT-PT
translate = {"日本": "Japão", 
             "中国": "China",
             "ประเทศไทย": "Tailândia",
             "Deutschland": "Alemanha",
             "España": "Espanha", 
             "Nederland": "Holanda",
             "New Zealand / Aotearoa": "Nova Zelândia",  
             "Schweiz/Suisse/Svizzera/Svizra": "Suíça", 
             "RP": "Polónia",
             "UK": "Reino Unido",
             "USA": "Estados Unidos da América"}
     
places = {}

for label in Document:
    if label.tag[32:] == "Folder":
        for place in label:    
            if place.tag[32:] == "Placemark":
                for attribute in place:
                    if attribute.tag[32:] == "name":
                        name = attribute.text
                        
                    elif attribute.tag[32:] == "Point":
                        coordinates = eval(attribute[0].text.strip()[:-2])[::-1]
                        place = geolocator.reverse(coordinates).raw                        
                        country = place["address"]["country"]
                        
                        if country in translate:
                            country = translate[country]
                        
                places[name] = country
                        
import xlsxwriter

# Create an new Excel file and add a worksheet.
workbook = xlsxwriter.Workbook('places.xlsx')
worksheet = workbook.add_worksheet()

# Widen the first column to make the text clearer.
worksheet.set_column("A:A", 23)
worksheet.set_column("B:B", 12)
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
    worksheet.write(counter, 0, places[place])
    worksheet.write(counter, 2, place)
    
    counter += 1

workbook.close()

end = time.time() 

print("Finished with success, it toke " + str(end - start) + " seconds to finish.")
    