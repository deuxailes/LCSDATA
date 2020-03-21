from urllib.request import urlopen as uReq
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font, colors
from openpyxl.formatting.rule import ColorScale, FormatObject, CellIsRule, ColorScaleRule
from decimal import *
import pandas
import json
from selenium import webdriver  # Import module
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import *
import time  # Waiting function

# This enables floats to be manipulated by the decimal library
wb = Workbook()
ws = wb.active

browser = webdriver.Firefox()
wait = WebDriverWait(browser, 10)
# Define URL
uClient = uReq("https://lol.gamepedia.com/LCS/2020_Season/Spring_Season/Scoreboards/Week_7")
matchHistoryPage = []
page_html = uClient.read()
page_soup = BeautifulSoup(page_html, "html.parser")

list = page_soup.findAll("div", attrs={"class": "inline-content"})
masterArray = []
playerLength = 0
minGold = 0
dummyDict = {}
day1Dict = {}
day2Dict = {}
day3Dict = {}
weekDict = {'DAY1': day1Dict, 'DAY2': day2Dict, 'DAY3': day3Dict}
positionArray = ['TOP', 'JUNG', 'MID', 'ADC', 'SUP']

def in_list(item, L):
    if item is None:
        return -1
    else:
        for x in L:
            if item in x:
                return L.index(x)
        return -1


def add_k(sheet):
    for row in sheet.iter_rows(min_row=2, min_col=2):
        for cell in row:
            if cell.value is not None:
                if float(cell.value):
                    cell.value = str(cell.value) + "k "


def format_color_cells(sheet):
    sheet.conditional_formatting.add('B1:B52',
                                     ColorScaleRule(start_type='percentile', start_value=10, start_color='ea7d7d',
                                                    mid_type='percentile', mid_value=50, mid_color='C0C0C0',
                                                    end_type='percentile', end_value=90, end_color='9de7b1'))

    sheet.conditional_formatting.add('D1:D52',
                                     ColorScaleRule(start_type='percentile', start_value=10, start_color='ea7d7d',
                                                    mid_type='percentile', mid_value=50, mid_color='C0C0C0',
                                                    end_type='percentile', end_value=90, end_color='9de7b1'))
    sheet.conditional_formatting.add('F1:F52',
                                     ColorScaleRule(start_type='percentile', start_value=10, start_color='AA0000',
                                                    mid_type='percentile', mid_value=50, mid_color='C0C0C0',
                                                    end_type='percentile', end_value=90, end_color='00AA00'))


def is_loaded(path):
    # path example: //a[@class="login-button btn-large btn-large-primary"]
    try:
        e = browser.find_element(By.XPATH, path)
    except NoSuchElementException:
        print("not loaded yet")
        e = None
    while e is None:
        try:
            e = browser.find_element(By.XPATH, path)
        except NoSuchElementException:
            print("not loaded yet")
            e = None
    return e


def main():
    for t in list:
        matchLink = t.find_all("div", {"class": "sb-datetime-mh"})
        for div in matchLink:
            i = div.find('a')
            matchHistoryPage.append(i['href'])

    browser.get('https://matchhistory.na.leagueoflegends.com/en/#page/landing-page')

    is_loaded('//a[@class="login-button btn-large btn-large-primary"]').click()
    is_loaded('//input[@name="username"]').send_keys('dryrhino4419')
    is_loaded('//input[@name="password"]').send_keys('gabythebaby1')
    time.sleep(7)
    browser.find_element_by_name("password").send_keys(Keys.ENTER)
    time.sleep(7)

    for i in range(len(matchHistoryPage)):
        browser.get(matchHistoryPage[i])
        time.sleep(3)
        element = wait.until(EC.presence_of_element_located((By.XPATH, '//div[@class="classic details"]')))
        page_html = element.get_attribute('innerHTML')
        page_soup = BeautifulSoup(page_html, "html.parser")

        playersRow = page_soup.findAll('div', {"class": "classic player"})
        j = 0
        teamDict = {}
        playerDict = {}
        k = 0

        statTable = page_soup.find('table', {"class": "table table-bordered"})
        # poop = statTable.findAll('tr')
        allTheStats = {}

        for x in range(len(statTable.contents[1].contents)):
            row = statTable.contents[1].contents[x]
            allTheStats.update({row.get_text(): row})

        minionsKilledRow = statTable.contents[1].contents[32]
        neutralMinionsKilledRow = statTable.contents[1].contents[33]
        neutralKilledInTEAMJRow = statTable.contents[1].contents[34]
        neutralKilledInENMJRow = statTable.contents[1].contents[35]
        largestKillingSpreeRow = statTable.contents[1].contents[3]

        for player in playersRow:
            statDick = {}
            playerGold = player.find('div', {"class": "gold-col gold"}).text[:-1]
            teamName = player.find('div', {"class": "champion-nameplate-name"}).get_text().split(" ", 3)[1]
            playerName = player.find('div', {"class": "champion-nameplate-name"}).get_text().split(" ", 3)[2]
            kills = player.find('div', {"class": "kda-kda"}).get_text().split("/", 3)[0]
            deaths = player.find('div', {"class": "kda-kda"}).get_text().split("/", 3)[1]
            assists = player.find('div', {"class": "kda-kda"}).get_text().split("/", 3)[2]
            champion = player.find('div', {"data-rg-name": "champion_10.4.1"}).get('data-rg-id')
            minionsKilled = minionsKilledRow.contents[j + 1].get_text()
            neutralMinionsKilled = neutralMinionsKilledRow.contents[j + 1].get_text()
            neutralKilledInTEAMJ = neutralKilledInTEAMJRow.contents[j + 1].get_text()
            neutralKilledInENMJ = neutralKilledInENMJRow.contents[j + 1].get_text()
            largestKillingSpree = largestKillingSpreeRow.contents[j + 1].get_text()

            if j == 5:
                playerDict = {}
                statDick = {}
                k = 0

            statDick.update({"index": k})
            statDick.update({"position": positionArray[k]})
            statDick.update({"champion": champion})
            statDick.update({"gold": playerGold})
            statDick.update({"kills": kills})
            statDick.update({"deaths": deaths})
            statDick.update({"assists": assists})
            statDick.update({"minions_killed": minionsKilled})
            statDick.update({"neutral_minions_killed": neutralMinionsKilled})
            statDick.update({"neutral_minions_killed_team_jungle": neutralKilledInTEAMJ})
            statDick.update({"neutral_minions_killed_enemy_jungle": neutralKilledInENMJ})
            statDick.update({"largest_killing_spree": largestKillingSpree})

            playerDict.update({playerName: statDick})

            k += 1
            if j > 5:
                teamName = player.find('div', {"class": "champion-nameplate-name"}).get_text().split(" ", 3)[1]

            teamDict.update({teamName: playerDict})

            j += 1

        time.sleep(2)

        if i < 4:
            weekDict['DAY1'].update(teamDict)
        elif i == 4 or i < 8:
            weekDict['DAY2'].update(teamDict)
        else:
            weekDict['DAY3'].update(teamDict)

    with open('player_info.JSON', 'w', encoding='utf-8') as f:
        json.dump(weekDict, f, ensure_ascii=False, indent=4)

    print(json.dump(allTheStats, f, ensure_ascii=False, indent=4))

    '''
    for i in range(len(masterArray)):
        row = i + 2
        ws.cell(row=row, column=1).value = masterArray[i][0]
        ws.cell(row=row, column=2).value = str(masterArray[i][1])
        if len(masterArray[i]) == 3:
            ws.cell(row=row, column=4).value = masterArray[i][2]
            ws.cell(row=row, column=6).value = Decimal((masterArray[i][2] + masterArray[i][1]) / Decimal(2))
    
    format_color_cells(ws)
    add_k(ws)

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
'''


main()

