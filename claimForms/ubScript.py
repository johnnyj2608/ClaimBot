from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException
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
        break   # FOR TESTING

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
        results = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(('xpath', '//*[@id="ui-id-1"]/li'))
        )
        if results.text == 'No results found':
            raise Exception
        options = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located(('xpath', '//*[@id="ui-id-1"]/li/a'))
        )
        if not options:
            raise Exception

        index = 0
        while True:
            patientsField.send_keys(Keys.ARROW_DOWN)
            if memberSelect == options[index].text:
                patientsField.send_keys(Keys.RETURN)
                break
            index += 1

    except (IndexError, TimeoutException, Exception):
        print(f"Element with visible text '{memberSelect}' not found.")
        patientsField.clear()
        return False
    
    payerCombo = Select(WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_PayerId"]'))
    ))
    billingProviderCombo = Select(WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ProviderId"]'))
    ))
    templatesCombo = Select(WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_TemplateId"]'))
    ))    
    physiciansCombo = Select(WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_AttendingPhysicianId"]'))
    ))
    createClaimButton = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="btnCreateClaim"]'))
    )

    templatesCombo.select_by_visible_text(insurance)

    createClaimButton.click()
    return True
    
def ubForm(driver, dxCode, authID, dates, autoSubmit, stopFlag):
    pass