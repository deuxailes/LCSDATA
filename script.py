from urllib.request import urlopen as uReq
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
from openpyxl.styles import colors
from openpyxl.styles import Font, Color
from decimal import *
import json

# This enables floats to be manipulated by the decimal library
getcontext()
Context(prec=28, rounding=ROUND_HALF_EVEN, Emin=-999999, Emax=999999,
        capitals=1, clamp=0, flags=[], traps=[Overflow, DivisionByZero,
        InvalidOperation])
getcontext().prec = 3

wb = Workbook()
ws = wb.active

uClient = uReq("https://lol.gamepedia.com/LCS/2020_Season/Spring_Season/Scoreboards/Week_7")
page_html = uClient.read()

# dataURL = "http://na.lolesports.com/api/player/84.json"
# response = uReq(dataURL)
# data = response.read().decode("UTF-8")
uClient.close()


page_soup = BeautifulSoup(page_html, "html.parser")

ws.cell(row=1, column=1).value = "Name"
ws.cell(row=1, column=2).value = "GD1"
ws.cell(row=1, column=3).value = "GD2"
ws.cell(row=1, column=4).value = "AG"

ws.cell(row=1, column=1).font = Font(size=14)
ws.cell(row=1, column=2).font = Font(size=14)
ws.cell(row=1, column=3).font = Font(size=14)
ws.cell(row=1, column=4).font = Font(size=14)

list = page_soup.findAll("div", attrs={"class": "inline-content"})

masterArray = []
playerLength = 0


def in_list(item, L):
    for x in L:
        if item in x:
            return L.index(x)
    return -1


for table in list:  # Sorts through tables
    # Sorts through gold count per player in given table
    tableGold = table.find_all("div", {"class": "sb-p-stat sb-p-stat-gold"})
    # Sorts through name per player in given table
    tableName = table.find_all("div", {"class": "sb-p-name"})
    for i in range(10):  # Max 10 players per match
        if in_list(tableName[i].text, masterArray) != -1:
            value_index = in_list(tableName[i].text, masterArray)
            masterArray[value_index].append(Decimal(tableGold[i].text[:-1]))
        else:
            masterArray.append([tableName[i].text, Decimal(tableGold[i].text[:-1])])

for i in range(len(masterArray)):
    row = i + 2
    ws.cell(row=row, column=1).value = masterArray[i][0]
    ws.cell(row=row, column=2).value = str(masterArray[i][1]) + "k "
    if len(masterArray[i]) == 3:
        ws.cell(row=row, column=3).value =  str(masterArray[i][2]) + "k "
        ws.cell(row=row, column=4).value =  str((masterArray[i][2] + masterArray[i][1]) / Decimal(2)) + "k "

# This adjusts cell width to biggest cell in column.
dims = {}
for row in ws.rows:
    for cell in row:
        if cell.value:
            dims[cell.column_letter] = max((dims.get(cell.column_letter, 0), len(str(cell.value))))
for col, value in dims.items():
    ws.column_dimensions[col].width = value

wb.save("lwt example.xlsx")
