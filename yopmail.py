from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import *
from time import sleep
from openpyxl import Workbook
import requests
import os
from datetime import datetime

driver = webdriver.Chrome()
driver.minimize_window()

workbook = Workbook()
sheet = workbook.active

url = 'https://yopmail.com/email-generator'

driver.get(url)
# newButton = driver.find_element(By.CLASS_NAME, 'nw').find_element(By.TAG_NAME, 'button')
# newButton.click()
username = driver.find_element(By.CLASS_NAME, 'segen').text;
driver.switch_to.frame('ifdoms')
domains = driver.find_elements(By.TAG_NAME, 'option')
for item in domains:
    domain = item.text
    if '@' in domain:
        sheet.append([username + domain])

# Format the date and time as per your requirement
current_datetime = datetime.now().strftime("%m%d%H%M")

workbook.save(os.path.normpath("./" + current_datetime + ".xlsx"))

print('Successfully saved!')
driver.quit()
