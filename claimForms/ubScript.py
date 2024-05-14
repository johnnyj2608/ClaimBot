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
        # memberName = lastName+', '+firstName+' ['+birthDate.strftime("%#m/%#d/%y")+']'
        memberName = firstName+' '+lastName

        if ubStored(driver, insurance, summary, memberName):
            dates = getDatesFromWeekdays(start, end, schedule, authStart, authEnd)
            total = ubForm(driver, dxCode, authID, dates, autoSubmit, stopFlag)

        completedMembers += 1
        statusLabel.configure(text=f"Completed Members: {completedMembers}/{totalMembers}")
        statusLabel.update()
        break # REMOVE AFTER TESTING

def ubStored(driver, insurance, summary, memberName):
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
        # Search Test McGee. Get all autofills. Select with correct bday
        patientsField.send_keys(memberName)
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(('xpath', '//*[@id="ui-id-1"]'))
        )
        patientsField.send_keys(Keys.ARROW_DOWN)
        patientsField.send_keys(Keys.RETURN)
        # Check if birth date is correct

    except NoSuchElementException:
        print(f"Element with visible text '{memberName}' not found.")
        return False
    
def ubForm(driver, dxCode, authID, dates, autoSubmit, stopFlag):
    pass