from urllib.request import urlopen as uReq
from bs4 import BeautifulSoup
import json
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait

url = "https://www.trackingthepros.com/players/"
page = uReq(url)
soup = BeautifulSoup(page, 'html.parser')
pageHTML = page.read()

driver = webdriver.Firefox()
driver.get(url)
submit_assignment = driver.find_element_by_id('form-control input-sm')
submit_assignent.click()
