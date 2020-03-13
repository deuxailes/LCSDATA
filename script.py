from urllib.request import urlopen as uReq
from bs4 import BeautifulSoup
import xlwt
from xlwt import Workbook

wb = Workbook()
sheet1 = wb.add_sheet('Sheet 1')
uClient = uReq("https://lol.gamepedia.com/LCS/2020_Season/Spring_Season/Scoreboards/Week_7")
page_html = uClient.read()
uClient.close()

page_soup = BeautifulSoup(page_html, "html.parser")


list = page_soup.findAll("div", attrs={"class":"inline-content"})

playerLength = 0
for table in list:
    tableGold = table.find_all("div", {"class": "sb-p-stat sb-p-stat-gold"})
    tableName = table.find_all("div", {"class": "sb-p-name"})
    for i in range(10):
        sheet1.write(playerLength + i + 1, 0, tableName[i].text)
        sheet1.write(playerLength + i + 1, 1, tableGold[i].text)
    playerLength += 10

wb.save("lwt example.xls")