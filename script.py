from urllib.request import urlopen as uReq
from bs4 import BeautifulSoup

url = input("Please enter website. \n")

uClient = uReq(url)
page_html = uClient.read()
uClient.close()

page_soup = soup(page_html, "html.parser")

