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
        if stopProcess(stopFlag): return

        lastName, firstName, birthDate, authID, dxCode, schedule, authStart, authEnd = member
        memberSearch = firstName+' '+lastName
        memberSelect = lastName+', '+firstName+' ['+birthDate.strftime("%#m/%#d/%y")+']'
        print(memberSearch, stopFlag)
        if ubStored(driver, insurance, summary, memberSearch, memberSelect):
            dates = getDatesFromWeekdays(start, end, schedule, authStart, authEnd)
            dates = intersectVacations(dates, start, end)
            total = ubForm(driver, dxCode, authID, start, end, dates, autoSubmit, stopFlag)

        completedMembers += 1
        statusLabel.configure(text=f"Completed Members: {completedMembers}/{totalMembers}")
        statusLabel.update()

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
    if stopProcess(stopFlag): return
    ub04URL = driver.current_url
    driver.switch_to.default_content()
    iframe = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="frame1"]'))
    )
    driver.switch_to.frame(iframe)

    memberIDField = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucUBForm_COBSubID1"]')))
    memberID = memberIDField.get_attribute("value")
    memberIDField = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucUBForm_PatientControlNumber"]')))
    memberIDField.send_keys(memberID)   # 3a
    memberIDField = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucUBForm_PatientID"]')))
    memberIDField.send_keys(memberID)   # 8a

    statementFromMonth = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucUBForm_StatementFromDate_Month"]')))
    statementFromDay = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucUBForm_StatementFromDate_Day"]')))
    statementFromYear = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucUBForm_StatementFromDate_Year"]')))
    
    statementToMonth = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucUBForm_StatementToDate_Month"]')))
    statementToDay = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucUBForm_StatementToDate_Day"]')))
    statementToYear = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucUBForm_StatementToDate_Year"]')))

    typeField = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucUBForm_TypeOfAdmission"]')))
    medicaidField = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucUBForm_COBPriorAuthNum1"]')))
    dxField = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucUBForm_PrimmaryDiagnosisCode"]')))
    
    statementFromMonth.send_keys(start.month)
    statementFromDay.send_keys(start.day)
    statementFromYear.send_keys(start.year)

    statementToMonth.send_keys(end.month)
    statementToDay.send_keys(end.day)
    statementToYear.send_keys(end.year)

    typeField.send_keys(9)
    dxField.clear()
    dxField.send_keys(dxCode)
    medicaidField.send_keys('N/A')

    addRowButton = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="divEdit2"]/div/table[6]/tbody/tr[101]/td[7]/i[1]')))

    # SDC -----

    revCodeDefault_S = driver.find_element(
            'xpath', '//*[@id="RevCode1"]')
    revCodeDefault_S = revCodeDefault_S.get_attribute("value")

    descDefault_S = driver.find_element(
            'xpath', '//*[@id="Description1"]')
    descDefault_S = descDefault_S.get_attribute("value")

    rateDefault_S = driver.find_element(
            'xpath', '//*[@id="Rate1"]')
    rateDefault_S = rateDefault_S.get_attribute("value")

    unitsDefault_S = driver.find_element(
            'xpath', '//*[@id="Units1"]')
    unitsDefault_S = unitsDefault_S.get_attribute("value")

    chargeDefault_S = driver.find_element(
            'xpath', '//*[@id="TotalCharge1"]')
    chargeDefault_S = chargeDefault_S.get_attribute("value")

    # TRANSPO -----

    revCodeDefault_T = driver.find_element(
            'xpath', '//*[@id="RevCode2"]')
    revCodeDefault_T = revCodeDefault_T.get_attribute("value")

    descDefault_T = driver.find_element(
            'xpath', '//*[@id="Description2"]')
    descDefault_T = descDefault_T.get_attribute("value")

    rateDefault_T = driver.find_element(
            'xpath', '//*[@id="Rate2"]')
    rateDefault_T = rateDefault_T.get_attribute("value")

    unitsDefault_T = driver.find_element(
            'xpath', '//*[@id="Units2"]')
    unitsDefault_T = unitsDefault_T.get_attribute("value")

    chargeDefault_T = driver.find_element(
            'xpath', '//*[@id="TotalCharge2"]')
    chargeDefault_T = chargeDefault_T.get_attribute("value")

    # If more than 12 * 2 dates
    for rowNum in range(23, (len(dates)*2)+1):
        if stopProcess(stopFlag): return
        addRowButton.click()

        revCodeRow = driver.find_element(
            'xpath', f'//*[@id="RevCode{rowNum}"]')

        descRow = driver.find_element(
            'xpath', f'//*[@id="Description{rowNum}"]')

        rateRow = driver.find_element(
            'xpath', f'//*[@id="Rate{rowNum}"]')

        unitsRow = driver.find_element(
            'xpath', f'//*[@id="Units{rowNum}"]')

        chargeRow = driver.find_element(
            'xpath', f'//*[@id="TotalCharge{rowNum}"]')
        if rowNum % 2 == 1:
            revCodeRow.send_keys(revCodeDefault_S)
            descRow.send_keys(descDefault_S)
            rateRow.send_keys(rateDefault_S)
            unitsRow.send_keys(unitsDefault_S)
            chargeRow.send_keys(chargeDefault_S)
        else:
            revCodeRow.send_keys(revCodeDefault_T)
            descRow.send_keys(descDefault_T)
            rateRow.send_keys(rateDefault_T)
            unitsRow.send_keys(unitsDefault_T)
            chargeRow.send_keys(chargeDefault_T)

    # If less than 12 * 2 dates
    for rowNum in range(22, len(dates)*2, -1):
        if stopProcess(stopFlag): return
        revCodeRow = driver.find_element(
                'xpath', f'//*[@id="RevCode{rowNum}"]')

        descRow = driver.find_element(
            'xpath', f'//*[@id="Description{rowNum}"]')

        rateRow = driver.find_element(
            'xpath', f'//*[@id="Rate{rowNum}"]')

        unitsRow = driver.find_element(
            'xpath', f'//*[@id="Units{rowNum}"]')

        chargeRow = driver.find_element(
            'xpath', f'//*[@id="TotalCharge{rowNum}"]')
        
        revCodeRow.clear()
        descRow.clear()
        rateRow.clear()
        unitsRow.clear()
        chargeRow.clear()

    for rowNum in range(len(dates)):
        if stopProcess(stopFlag): return
        curDate = dates[rowNum]
        month = curDate.month
        day = curDate.day
        year = curDate.year

        rowNum = (rowNum * 2) + 1    # Index 1
        for i in range(2):
            rowNum += i
            fmMonth = driver.find_element(
                'xpath', f'//*[@id="FromDateMonth{rowNum}"]')
            
            fmDay = driver.find_element(
                'xpath', f'//*[@id="FromDateDay{rowNum}"]')
            
            fmYear = driver.find_element(
                'xpath', f'//*[@id="FromDateYear{rowNum}"]')
            
            toMonth = driver.find_element(
                'xpath', f'//*[@id="ToDateMonth{rowNum}"]')
            
            toDay = driver.find_element(
                'xpath', f'//*[@id="ToDateDay{rowNum}"]')
            
            toYear = driver.find_element(
                'xpath', f'//*[@id="ToDateYear{rowNum}"]')
            
            fmMonth.send_keys(month)
            fmDay.send_keys(day)
            fmYear.send_keys(year)
            toMonth.send_keys(month)
            toDay.send_keys(day)
            toYear.send_keys(year)

    removeRowButton = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="divEdit2"]/div/table[6]/tbody/tr[101]/td[7]/i[2]')))
    removeRowButton.click()

    totalField = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucUBForm_TotalCharge"]')))
    total = totalField.get_attribute("value")

    if stopProcess(stopFlag): return
    if autoSubmit:
        submitButton = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucUBForm_btnSCUpdate"]')))
        submitButton.click()
        # Wait to see next page before returning
    else:
        while driver.current_url == ub04URL:
            time.sleep(1)
    return total