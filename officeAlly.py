
from datetime import datetime, timedelta

def getDatesFromWeekdays(month, year, weekdays):
    numDays = (datetime(year, month % 12 + 1, 1) - timedelta(days=1)).day
    monthDays = []

    for day in range(1, numDays+1):
        date = datetime(year, month, day)
        if date.weekday() in weekdays:
            monthDays.append(date.strftime("%m/%d/%Y"))

    return monthDays

dates = getDatesFromWeekdays(4,2024,[0, 1, 2, 3, 4])

office_ally = 'https://www.officeally.com/secure_oa.asp'

from selenium import webdriver

options = webdriver.ChromeOptions()
options.add_experimental_option("detach", True)

driver = webdriver.Chrome(options=options)
driver.get(office_ally)
driver.maximize_window()
driver.implicitly_wait(10)

username = 'username'
password = 'password'

username_field = driver.find_element('xpath', '//*[@id="username"]')
username_field.send_keys(username)

password_field = driver.find_element('xpath', '//*[@id="password"]')
password_field.send_keys(password)

login_button = driver.find_element('xpath', '/html/body/main/section/div/div/div/form/div[2]/button')
login_button.click()

driver.get('https://www.officeally.com/secure_oa.asp?GOTO=OnlineEntry&TaskAction=Manage')

iframe = driver.find_element('xpath', '//*[@id="Iframe9"]')
driver.switch_to.frame(iframe)

create_claim_button = driver.find_element('xpath', '//*[@id="Button2"]')
create_claim_button.click()

driver.switch_to.default_content()
iframe = driver.find_element('xpath', '//*[@id="Iframe9"]')
driver.switch_to.frame(iframe)

add_row_button = driver.find_element('xpath', '//*[@id="btnAddRow"]')
for i in range(len(dates)-12):
    add_row_button.click()