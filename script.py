from urllib.request import urlopen as uReq
from bs4 import BeautifulSoup
from openpyxl import Workbook
import pandas as pd


wb = Workbook()
ws = wb.active


uClient = uReq("https://lol.gamepedia.com/LCS/2020_Season/Spring_Season/Scoreboards/Week_7")
page_html = uClient.read()
uClient.close()

page_soup = BeautifulSoup(page_html, "html.parser")

ws.cell(row=1, column=1).value = "Name"
ws.cell(row=1, column=2).value = "GD1"
ws.cell(row=1, column=3).value = "GD2"
ws.cell(row=1, column=4).value = "GD3"
ws.cell(row=1, column=5).value = "AG"

list = page_soup.findAll("div", attrs={"class": "inline-content"})

masterArray= []
playerLength = 0

for table in list: # Sorts through tables
    # Sorts through gold count per player in given table
    tableGold = table.find_all("div", {"class": "sb-p-stat sb-p-stat-gold"})
    # Sorts through name per player in given table
    tableName = table.find_all("div", {"class": "sb-p-name"})
    for i in range(10): # Max 10 players per match
        if tableName[i].text in masterArray:
            value_index = masterArray.index(tableName[i].text)
            if masterArray[value_index][1]:
                masterArray[value_index][2] = tableGold[i].text
            else:
                masterArray[value_index][1] = tableGold[i].text
        else:
            masterArray.append([tableName[i].text, tableGold[i].text])

 #   playerLength += 10

print(masterArray)
#ws.cell(row=playerLength + i + 2, column=1).value = tableName[i].text
#ws.cell(row=playerLength + i + 2, column=2).value = tableGold[i].text
#wb.save("lwt example.xlsx")
