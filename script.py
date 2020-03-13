from urllib.request import urlopen as uReq
from bs4 import BeautifulSoup



uClient = uReq("https://lol.gamepedia.com/LCS/2020_Season/Spring_Season/Scoreboards/Week_7")
page_html = uClient.read()
uClient.close()

page_soup = BeautifulSoup(page_html, "html.parser")


list = page_soup.findAll("div", attrs={"class":"inline-content"})

for table in list:
    tableGold = table.find_all("div", {"class": "sb-p-stat sb-p-stat-gold"})
    tableName = table.find_all("div", {"class": "sb-p-name"})
    print(tableGold.text)
    print(tableName.text)

