from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font, colors
from openpyxl.formatting.rule import ColorScale, FormatObject, CellIsRule, ColorScaleRule
from openpyxl.utils import get_column_letter
from decimal import *
import pandas
import json

wb = Workbook()
ws = wb.active
redFill = PatternFill(start_color='EE1111', end_color='EE1111', fill_type='solid')
getcontext()
Context(prec=28, rounding=ROUND_HALF_EVEN, Emin=-999999, Emax=999999,
        capitals=1, clamp=0, flags=[], traps=[Overflow, DivisionByZero,
                                              InvalidOperation])
getcontext().prec = 4

with open('player_info.JSON', 'r') as myfile:
    data = myfile.read()

obj = json.loads(data)


def initial_formatting():
    ws.row_dimensions[1].height = 25.5
    ws.row_dimensions[2].height = 21.75
    ws.row_dimensions[3].height = 29.25
    ws.column_dimensions['K'].width = .5
    ws.merge_cells('A1:J2')
    ws.merge_cells('L1:V1')
    ws.cell(row=3, column=1).value = "Name"
    ws.cell(row=3, column=2).value = "GD1  "
    ws.cell(row=3, column=3).value = "MP1  "
    ws.cell(row=3, column=4).value = "GD2  "
    ws.cell(row=3, column=5).value = "MP2  "
    ws.cell(row=3, column=6).value = "AvG  "
    ws.cell(row=3, column=7).value = "AMP"
    ws.cell(row=3, column=8).value = "GPM"

    ws.cell(row=3, column=1).font = Font(size=14)
    ws.cell(row=3, column=2).font = Font(size=12)
    ws.cell(row=3, column=3).font = Font(size=12)
    ws.cell(row=3, column=4).font = Font(size=12)
    ws.cell(row=3, column=5).font = Font(size=12)
    ws.cell(row=3, column=6).font = Font(size=12)
    ws.cell(row=3, column=7).font = Font(size=12)
    ws.cell(row=3, column=8).font = Font(size=12)


def post_formatting():
    column_widths = []
    for row in ws.iter_rows(3, 43):
        for i, cell in enumerate(row):
            try:
                column_widths[i] = max(column_widths[i], len(str(cell.value)))
            except IndexError:
                try:
                    column_widths.append(len(cell.value))
                except TypeError:
                    return

    for i, column_width in enumerate(column_widths):
        ws.column_dimensions[get_column_letter(i + 1)].width = column_width + 1


def main():
    initial_formatting()
    j = 0
    k = 0
    row = 4
    exists = False
    first_column = ws['A']

    for day in obj:
        for team in obj[day]:
            for player in obj[day][team]:
                first_column = ws['A']
                for x in range(len(first_column)): # Checks if player exists
                    if first_column[x].value == player:
                        ws.cell(row=x+1, column=4).value = obj[day][team][player]['gold']
                        exists = True
                if exists is False:
                    ws.cell(row=row, column=1).value = player
                    ws.cell(row=row, column=2).value = obj[day][team][player]['gold']
                    row += 1
                exists = False

    post_formatting()
    wb.save("lwt example.xlsx")


main()
