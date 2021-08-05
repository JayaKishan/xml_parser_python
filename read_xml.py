from bs4 import BeautifulSoup
import xml.etree.ElementTree as ET
import os
from pathlib import Path
import xlsxwriter
import smtplib, ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


path = 'F:/Learnings/Python/Project/xml_files/'
output_path = 'F:/Learnings/Python/Project/output/'

workbook = xlsxwriter.Workbook(output_path+'report.xlsx')
worksheet = workbook.add_worksheet()
heading_cell_format = workbook.add_format({'bold': True, 'font_color': 'black','font_size': 15,'bg_color': 'yellow'})

worksheet.write('A1', 'Product Names',heading_cell_format)
worksheet.write('B1', 'Country Code',heading_cell_format)
worksheet.write('C1', 'Country Name',heading_cell_format)
worksheet.write('D1', 'Route of administration',heading_cell_format)
worksheet.write('E1', 'Formulation',heading_cell_format)
worksheet.write('F1', 'Event Start Date',heading_cell_format)
worksheet.write('G1', 'Event End Date',heading_cell_format)
worksheet.write('H1', 'XML File Name',heading_cell_format)

row = 1
column = 0

xml_name_row = 1
dis_name_row = 1
country_code_row = 1
start_date_row = 1
end_date_row = 1

for filename in os.listdir(path):
    #worksheet.write(row, column+7, filename)
    #xml_name_row+=1


    if not filename.endswith('.xml'): continue
    fullname = os.path.join(path, filename)
    tree = ET.parse(fullname)
    
    with open(fullname, 'r') as f:
        data = f.read()

    xml_data = BeautifulSoup(data, "xml")
    names = xml_data.find_all('name')
    attrvalues = xml_data.find_all("routeCode")
    #locatedPlace = xml_data.find_all("locatedPlace")
    countries   = xml_data.find_all("code", {"codeSystem": "1.0.3166.1.2.2"})
    start_dates = xml_data.find_all("observation")
    #end_dates   = xml_data.find_all("effectiveTime", {"xsi:type": "IVL_TS"})

    for name in names:
        if len(name.text) != 0:
            worksheet.write(xml_name_row, column, name.text)
            worksheet.write(xml_name_row, column+7, filename)
            xml_name_row+=1
        
    #for start_date in start_dates:
    #    print(start_date, start_date.attrib)
        #if  ( start_date.get("code") != None  and start_date.get("displayName") != None ):
        #if start_date.find("code",{"displayName","reaction"}) 
        #if start_date.findChildren() != None:
        #    print(start_date)
        #if start_date.findChild() == "low":
        #    child_rec = start_date.findChild()
        #    print(child_rec)
        #    worksheet.write(start_date_row, column+5, child_rec["value"])
        #    start_date_row+=1
    #for end_date in end_dates:
    #    if len(end_date["value"]) != 0:
    #        worksheet.write(end_date_row, column+6, end_date["value"])
    #        end_date_row +=1
    
    for country in countries:
        if  ( country.get("code") != None  and country.get("displayName") != None ):
            worksheet.write(country_code_row, column+1, country["code"])
            worksheet.write(country_code_row, column+2, country["displayName"])
            country_code_row+=1
    
    for attrvalue in attrvalues:
        if len(attrvalue['displayName']) != 0:
            worksheet.write(dis_name_row, column+3, attrvalue['displayName'])
            dis_name_row+=1
workbook.close()


