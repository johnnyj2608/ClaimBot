from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.keys import Keys
from claimForms.claimFormsHelper import *
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
            dates = intersectVacations(dates, start, end)
            total = ubForm(driver, dxCode, authID, start, end, dates, autoSubmit, stopFlag)

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
    
def ubForm(driver, dxCode, authID, start, end, dates, autoSubmit, stopFlag):
    ub04URL = driver.current_url
    driver.switch_to.default_content()
    iframe = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="frame1"]'))
    )
    driver.switch_to.frame(iframe)

    # member ID field: copy from 60. paste 3a & 8a
    # medicaid ID
    # 14 type 9
    # 67 dx
    # 6 statement period
    # 63 medicaid

    memberIDField = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucUBForm_COBSubID1"]')))
    memberID = memberIDField.get_attribute("value")
    memberIDField = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucUBForm_PatientControlNumber"]')))
    memberIDField.send_keys(memberID)
    memberIDField = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucUBForm_PatientID"]')))
    memberIDField.send_keys(memberID)

    authField = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucHCFA_PRIOR_AUTH_NUMBER"]')))
    authField.send_keys(authID)

    dxField = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucHCFA_DIAGNOSIS_CODECMS0212_1"]')))
    dxField.clear()
    dxField.send_keys(dxCode)

    addRowButton = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="btnAddRow"]')))

    placeDefault = driver.find_element(
            'xpath', 
            '//*[@id="ctl00_phFolderContent_ucHCFA_ucHCFALineItem_ucClaimLineItem_PLACE_OF_SVC0"]')
    placeDefault = placeDefault.get_attribute("value")

    cptDefault = driver.find_element(
            'xpath', 
            '//*[@id="ctl00_phFolderContent_ucHCFA_ucHCFALineItem_ucClaimLineItem_CPT_CODE0"]')
    cptDefault = cptDefault.get_attribute("value")

    modifierDefault = driver.find_element(
            'xpath', 
            '//*[@id="ctl00_phFolderContent_ucHCFA_ucHCFALineItem_ucClaimLineItem_MODIFIER_A0"]')
    modifierDefault = modifierDefault.get_attribute("value")

    diagnosisDefault = driver.find_element(
            'xpath', 
            '//*[@id="ctl00_phFolderContent_ucHCFA_ucHCFALineItem_ucClaimLineItem_DOS_DIAG_CODE0"]')
    diagnosisDefault = diagnosisDefault.get_attribute("value")

    chargeDefault = driver.find_element(
            'xpath', 
            '//*[@id="ctl00_phFolderContent_ucHCFA_ucHCFALineItem_ucClaimLineItem_DOS_CHRG0"]')
    chargeDefault = chargeDefault.get_attribute("value")

    unitsDefault = driver.find_element(
            'xpath', 
            '//*[@id="ctl00_phFolderContent_ucHCFA_ucHCFALineItem_ucClaimLineItem_UNITS0"]')
    unitsDefault = unitsDefault.get_attribute("value")

    if len(dates) > 12:
        for rowNum in range(12, len(dates)):
            addRowButton.click()

            placeRow = driver.find_element(
                'xpath', 
                f'//*[@id="ctl00_phFolderContent_ucHCFA_ucHCFALineItem_ucClaimLineItem_PLACE_OF_SVC{rowNum}"]')
            
            cptRow = driver.find_element(
                'xpath', 
                f'//*[@id="ctl00_phFolderContent_ucHCFA_ucHCFALineItem_ucClaimLineItem_CPT_CODE{rowNum}"]')
            
            modifierRow = driver.find_element(
                'xpath', 
                f'//*[@id="ctl00_phFolderContent_ucHCFA_ucHCFALineItem_ucClaimLineItem_MODIFIER_A{rowNum}"]')
            
            diagnosisRow = driver.find_element(
                'xpath', 
                f'//*[@id="ctl00_phFolderContent_ucHCFA_ucHCFALineItem_ucClaimLineItem_DOS_DIAG_CODE{rowNum}"]')
            
            chargeRow = driver.find_element(
                'xpath', 
                f'//*[@id="ctl00_phFolderContent_ucHCFA_ucHCFALineItem_ucClaimLineItem_DOS_CHRG{rowNum}"]')
            
            unitsRow = driver.find_element(
                'xpath', 
                f'//*[@id="ctl00_phFolderContent_ucHCFA_ucHCFALineItem_ucClaimLineItem_UNITS{rowNum}"]')
            
            placeRow.send_keys(placeDefault)
            cptRow.send_keys(cptDefault)
            modifierRow.send_keys(modifierDefault)
            diagnosisRow.send_keys(diagnosisDefault)
            chargeRow.send_keys(chargeDefault)
            unitsRow.send_keys(unitsDefault)
    else:
        for rowNum in range(11, len(dates)-1, -1):
            placeRow = driver.find_element(
                'xpath', 
                f'//*[@id="ctl00_phFolderContent_ucHCFA_ucHCFALineItem_ucClaimLineItem_PLACE_OF_SVC{rowNum}"]')
            
            cptRow = driver.find_element(
                'xpath', 
                f'//*[@id="ctl00_phFolderContent_ucHCFA_ucHCFALineItem_ucClaimLineItem_CPT_CODE{rowNum}"]')
            
            modifierRow = driver.find_element(
                'xpath', 
                f'//*[@id="ctl00_phFolderContent_ucHCFA_ucHCFALineItem_ucClaimLineItem_MODIFIER_A{rowNum}"]')
            
            diagnosisRow = driver.find_element(
                'xpath', 
                f'//*[@id="ctl00_phFolderContent_ucHCFA_ucHCFALineItem_ucClaimLineItem_DOS_DIAG_CODE{rowNum}"]')
            
            chargeRow = driver.find_element(
                'xpath', 
                f'//*[@id="ctl00_phFolderContent_ucHCFA_ucHCFALineItem_ucClaimLineItem_DOS_CHRG{rowNum}"]')
            
            unitsRow = driver.find_element(
                'xpath', 
                f'//*[@id="ctl00_phFolderContent_ucHCFA_ucHCFALineItem_ucClaimLineItem_UNITS{rowNum}"]')
            
            placeRow.clear()
            cptRow.clear()
            modifierRow.clear()
            diagnosisRow.clear()
            chargeRow.clear()
            unitsRow.clear()

    for rowNum in range(len(dates)):
        curDate = dates[rowNum]
        month = curDate.month
        day = curDate.day
        year = curDate.year

        fmMonth = driver.find_element(
            'xpath', 
            f'//*[@id="ctl00_phFolderContent_ucHCFA_ucHCFALineItem_ucClaimLineItem_FM_DATE_OF_SVC_MONTH{rowNum}"]'
            )
        
        fmDay = driver.find_element(
            'xpath', 
            f'//*[@id="ctl00_phFolderContent_ucHCFA_ucHCFALineItem_ucClaimLineItem_FM_DATE_OF_SVC_DAY{rowNum}"]'
            )
        
        fmYear = driver.find_element(
            'xpath', 
            f'//*[@id="ctl00_phFolderContent_ucHCFA_ucHCFALineItem_ucClaimLineItem_FM_DATE_OF_SVC_YEAR{rowNum}"]'
            )
        
        toMonth = driver.find_element(
            'xpath', 
            f'//*[@id="ctl00_phFolderContent_ucHCFA_ucHCFALineItem_ucClaimLineItem_TO_DATE_OF_SVC_MONTH{rowNum}"]'
            )
        
        toDay = driver.find_element(
            'xpath', 
            f'//*[@id="ctl00_phFolderContent_ucHCFA_ucHCFALineItem_ucClaimLineItem_TO_DATE_OF_SVC_DAY{rowNum}"]'
            )
        
        toYear = driver.find_element(
            'xpath', 
            f'//*[@id="ctl00_phFolderContent_ucHCFA_ucHCFALineItem_ucClaimLineItem_TO_DATE_OF_SVC_YEAR{rowNum}"]'
            )
        
        fmMonth.send_keys(month)
        fmDay.send_keys(day)
        fmYear.send_keys(year)
        toMonth.send_keys(month)
        toDay.send_keys(day)
        toYear.send_keys(year)

    totalField = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucHCFA_TOTAL_CHARGE"]')))
    total = totalField.get_attribute("value")

    if stopFlag.value:
        stopFlag.value = False
    elif autoSubmit:
        submitButton = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucHCFA_btnSCUpdate"]')))
        submitButton.click()
    else:
        while driver.current_url == cms1500URL:
            time.sleep(1)
    return total