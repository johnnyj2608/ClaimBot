from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select

from datetime import datetime, timedelta

def getDatesFromWeekdays(month, year, weekdays):
    numDays = (datetime(year, month % 12 + 1, 1) - timedelta(days=1)).day
    monthDays = []

    for day in range(1, numDays+1):
        date = datetime(year, month, day)
        if date.weekday() in weekdays:
            monthDays.append(date.strftime("%m/%d/%Y"))

    return monthDays

def officeAllyAutomate(insurance, summary, members):
    office_ally = 'https://www.officeally.com/secure_oa.asp'

    options = webdriver.ChromeOptions()
    options.add_experimental_option("detach", True)

    driver = webdriver.Chrome(options=options)
    driver.get(office_ally)
    driver.maximize_window()
    # driver.implicitly_wait(10)

    # ------------- LOGIN PAGE ------------

    usernameField = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="username"]'))
    )
    passwordField = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="password"]'))
    )
    loginButton = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '/html/body/main/section/div/div/div/form/div[2]/button'))
    )

    usernameField.send_keys(summary['username'])
    passwordField.send_keys(summary['password'])
    loginButton.click()

    # ------------- SUMMARY PAGE ------------

    driver.get('https://www.officeally.com/secure_oa.asp?GOTO=OnlineEntry&TaskAction=Manage')

    iframe = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="Iframe9"]'))
    )
    driver.switch_to.frame(iframe)

    payerCombo = Select(WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ddlPayer"]'))
    ))
    patientsCombo = Select(WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ddlPatient"]'))
    ))
    billingProviderCombo = Select(WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ddlBillingProvider"]'))
    ))
    renderingProviderCombo = Select(WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ddlRenderingProvider"]'))
    ))
    facilitiesCombo = Select(WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ddlFacility"]'))
    ))
    templatesCombo = Select(WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ddlTemplate"]'))
    ))
    createClaimButton = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="Button2"]'))
    )

    renderingProviderCombo.select_by_visible_text(summary['renderingProvider'])
    facilitiesCombo.select_by_visible_text(summary['facilities'])
    templatesCombo.select_by_visible_text(insurance)

    # Loop starts here
    # lastName, firstName, [mm/dd/yyyy]
    # If patient not found, skip. Write log of events

    createClaimButton.click()

    # ------------- CLAIMS PAGE ------------

    # ------------- CMS-1500 ------------

    driver.switch_to.default_content()
    iframe = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="Iframe9"]'))
    )
    driver.switch_to.frame(iframe)

    addRowButton = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="btnAddRow"]'))
    )

    dates = getDatesFromWeekdays(4,2024,[0, 1, 2, 3, 4])
    for _ in range(len(dates)-12):
        addRowButton.click()

    # Insert member days data here

summaryValues = {
                "billingProvider": 'billingProvider',
                "renderingProvider": 'Provider, Rendering [123]',
                "facilities": 'facilities',
                "username": 'username',
                "password": 'password'
            }

officeAllyAutomate('Insurance', summaryValues, [])