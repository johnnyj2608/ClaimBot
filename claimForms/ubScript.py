from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
from scheduleParser import getDatesFromWeekdays
import time

def ubScript(driver, 
              insurance,
              summary, 
              members, 
              start, 
              end, 
              autoSubmit, 
              statusLabel, 
              stopFlag):
    
    totalMembers, completedMembers = len(members), 0
    for member in members:
        if stopFlag.value:
            stopFlag.value = False
            print("Automation stopped")
            break

        lastName, firstName, birthDate, authID, dxCode, schedule, authStart, authEnd = member
        memberSearch = firstName+' '+lastName
        memberSelect = lastName+', '+firstName+' ['+birthDate.strftime("%#m/%#d/%y")+']'

        if ubStored(driver, insurance, summary, memberSearch, memberSelect):
            dates = getDatesFromWeekdays(start, end, schedule, authStart, authEnd)
            total = ubForm(driver, dxCode, authID, dates, autoSubmit, stopFlag)

        completedMembers += 1
        statusLabel.configure(text=f"Completed Members: {completedMembers}/{totalMembers}")
        statusLabel.update()
        break # REMOVE AFTER TESTING

def ubStored(driver, insurance, summary, memberSearch, memberSelect):
    storedInfoURL='https://www.officeally.com/secure_oa.asp?GOTO=UB04OnlineEntry&TaskAction=StoredInfo'
    if driver.current_url != storedInfoURL:
        driver.get(storedInfoURL)
        driver.refresh()    # Required for auto-complete
        iframe = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(('xpath', '//*[@id="Iframe24"]'))
        )
        driver.switch_to.frame(iframe)

    patientsField = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_txtbxPatientId"]'))
    )

    try:
        patientsField.send_keys(memberSearch)
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(('xpath', '//*[@id="ui-id-1"]'))
        )
        options = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located(('xpath', '//*[@id="ui-id-1"]/li/a'))
        )

        index = 0
        while True:
            patientsField.send_keys(Keys.ARROW_DOWN)
            if memberSelect == options[index].text:
                patientsField.send_keys(Keys.RETURN)
                break
            index += 1
        print('found')

    except IndexError:
        print(f"Element with visible text '{memberSelect}' not found.")
        return False
    
def ubForm(driver, dxCode, authID, dates, autoSubmit, stopFlag):
    pass