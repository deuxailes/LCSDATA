from urllib.request import urlopen as uReq
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font, colors
from openpyxl.formatting.rule import ColorScale, FormatObject, CellIsRule, ColorScaleRule
from decimal import *
import json

# This enables floats to be manipulated by the decimal library
from openpyxl.utils import get_column_letter

getcontext()
Context(prec=28, rounding=ROUND_HALF_EVEN, Emin=-999999, Emax=999999,
        capitals=1, clamp=0, flags=[], traps=[Overflow, DivisionByZero,
                                              InvalidOperation])
getcontext().prec = 4

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
ws.cell(row=1, column=2).value = "GD1  "
ws.cell(row=1, column=3).value = "MP1  "
ws.cell(row=1, column=4).value = "GD2  "
ws.cell(row=1, column=5).value = "MP2  "
ws.cell(row=1, column=6).value = "AG   "
ws.cell(row=1, column=7).value = "MP"
ws.cell(row=1, column=8).value = "MP"

ws.cell(row=1, column=1).font = Font(size=14)
ws.cell(row=1, column=2).font = Font(size=12)
ws.cell(row=1, column=3).font = Font(size=12)
ws.cell(row=1, column=4).font = Font(size=12)
ws.cell(row=1, column=5).font = Font(size=12)
ws.cell(row=1, column=6).font = Font(size=12)
ws.cell(row=1, column=7).font = Font(size=12)
ws.cell(row=1, column=8).font = Font(size=12)

list = page_soup.findAll("div", attrs={"class": "inline-content"})


masterArray = []
playerLength = 0
redFill = PatternFill(start_color='EE1111', end_color='EE1111', fill_type='solid')
minGold = 0



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
    ws.cell(row=row, column=2).value = masterArray[i][1]
    if len(masterArray[i]) == 3:
        ws.cell(row=row, column=4).value = masterArray[i][2]
        ws.cell(row=row, column=6).value = Decimal((masterArray[i][2] + masterArray[i][1]) / Decimal(2))

ws.conditional_formatting.add('B1:B52',
                              ColorScaleRule(start_type='percentile', start_value=10, start_color='ea7d7d',
                                             mid_type='percentile', mid_value=50, mid_color='C0C0C0',
                                             end_type='percentile', end_value=90, end_color='9de7b1'))

ws.conditional_formatting.add('D1:D52',
                              ColorScaleRule(start_type='percentile', start_value=10, start_color='ea7d7d',
                                             mid_type='percentile', mid_value=50, mid_color='C0C0C0',
                                             end_type='percentile', end_value=90, end_color='9de7b1'))
ws.conditional_formatting.add('F1:F52',
                              ColorScaleRule(start_type='percentile', start_value=10, start_color='AA0000',
                                             mid_type='percentile', mid_value=50, mid_color='C0C0C0',
                                             end_type='percentile', end_value=90, end_color='00AA00'))



# This adjusts cell width to biggest cell in column.
column_widths = []
for row in ws.iter_rows():
    for i, cell in enumerate(row):
        try:
            column_widths[i] = max(column_widths[i], len(str(cell.value)))
        except IndexError:
            column_widths.append(len(cell.value))

for i, column_width in enumerate(column_widths):
    ws.column_dimensions[get_column_letter(i + 1)].width = column_width


wb.save("lwt example.xlsx")
