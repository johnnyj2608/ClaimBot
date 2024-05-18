from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
from claimForms.claimFormsHelper import *
from excel import recordClaims
import time
import glob
import os

def ubScript(driver, 
              summary, 
              members, 
              start, 
              end, 
              filePath,
              autoSubmit,
              autoDownload,
              statusLabel, 
              stopFlag):
    
    totalMembers, completedMembers = len(members), 0
    statusLabel.configure(text=f"Completed Members: {completedMembers}/{totalMembers}")
    statusLabel.update()
    
    summaryStats = {
        'members': totalMembers,
        'submitted': 0,
        'excluded': 0,
        'total': 0,
        'failed': 0
    }

    for member in members:
        if stopProcess(stopFlag): return

        if not member['exclude']:
            memberName = member['lastName']+', '+member['firstName']
            memberSearch = member['firstName']+' '+member['lastName']
            memberSelect = member['lastName']+', '+member['firstName']+' ['+member['birthDate'].strftime("%#m/%#d/%y")+']'

            total = -1
            if ubStored(driver, summary, memberSearch, memberSelect):
                dates = getDatesFromWeekdays(start, end, member['schedule'], member['authStart'], member['authEnd'])
                dates = intersectVacations(dates, member['vacationStart'], member['vacationEnd'])
                total = ubForm(driver, summary, member['dxCode'], member['medicaid'],
                            start, end, dates, autoSubmit, stopFlag)
            if stopProcess(stopFlag): return
            if total == -1:
                summaryStats['failed'] += 1
            else:
                summaryStats['submitted'] += 1
                ubDownload(driver, autoDownload, memberName, stopFlag)
            recordClaims(filePath, 
                         memberName,
                         start.strftime("%#m/%#d/%y")+' - '+end.strftime("%#m/%#d/%y"),
                         total)
        else:
            summaryStats['excluded'] += 1
        completedMembers += 1
        statusLabel.configure(text=f"Completed Members: {completedMembers}/{totalMembers}")
        statusLabel.update()
    return summaryStats

def ubStored(driver, summary, memberSearch, memberSelect):
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
        patientsField.clear()
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

    except (IndexError, Exception):
        print(f"Element with visible text '{memberSelect}' not found.")
        return False
    
    payerCombo = Select(WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_PayerId"]'))
    ))
    billingProviderCombo = Select(WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ProviderId"]'))
    )) 
    physiciansCombo = Select(WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_AttendingPhysicianId"]'))
    ))
    createClaimButton = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="btnCreateClaim"]'))
    )

    try:
        payerCombo.select_by_visible_text(summary['payer'])
        billingProviderCombo.select_by_visible_text(summary['billingProvider'])
        physiciansCombo.select_by_visible_text(summary['physician'])
    except NoSuchElementException:
        print("Failed to select stored information")
        return False

    createClaimButton.click()
    return True
    
def ubForm(driver, summary, dxCode, medicaid, start, end, dates, autoSubmit, stopFlag):
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
    statementFromMonth.send_keys(start.month)
    
    statementFromDay = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucUBForm_StatementFromDate_Day"]')))
    statementFromDay.send_keys(start.day)
    
    statementFromYear = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucUBForm_StatementFromDate_Year"]')))
    statementFromYear.send_keys(start.year)
    
    statementToMonth = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucUBForm_StatementToDate_Month"]')))
    statementToMonth.send_keys(end.month)
    
    statementToDay = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucUBForm_StatementToDate_Day"]')))
    statementToDay.send_keys(end.day)
    
    statementToYear = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucUBForm_StatementToDate_Year"]')))
    statementToYear.send_keys(end.year)

    billTypeField = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucUBForm_BillType"]')))
    billTypeField.send_keys(summary['billType'])
    
    admissionTypeField = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucUBForm_TypeOfAdmission"]')))
    admissionTypeField.send_keys(9)
    
    medicaidField = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucUBForm_COBPriorAuthNum1"]')))
    medicaidField.send_keys(medicaid)
    
    dxField = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucUBForm_PrimmaryDiagnosisCode"]')))
    dxField.send_keys(dxCode)
    
    for rowNum in range(1, (len(dates)*2)+1):
        if stopProcess(stopFlag): return

        revCodeRow = driver.find_element(
            'xpath', f'//*[@id="RevCode{rowNum}"]')
        revCodeRow.send_keys(summary['revenueCode'])

        descRow = driver.find_element(
            'xpath', f'//*[@id="Description{rowNum}"]')

        cptRow = driver.find_element(
            'xpath', f'//*[@id="Rate{rowNum}"]')

        unitsRow = driver.find_element(
            'xpath', f'//*[@id="Units{rowNum}"]')

        chargeRow = driver.find_element(
            'xpath', f'//*[@id="TotalCharge{rowNum}"]')
        
        if rowNum % 2 == 1:
            descRow.send_keys(summary['descriptionSDC'])
            cptRow.send_keys(summary['cptCodeSDC'])
            unitsRow.send_keys(summary['unitsSDC'])
            chargeRow.send_keys(summary['chargesSDC'])
        else:
            descRow.send_keys(summary['descriptionTrans'])
            cptRow.send_keys(summary['cptCodeTrans'])
            unitsRow.send_keys(summary['unitsTrans'])
            chargeRow.send_keys(summary['chargesTrans'])

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
            fmMonth.send_keys(month)
            
            fmDay = driver.find_element(
                'xpath', f'//*[@id="FromDateDay{rowNum}"]')
            fmDay.send_keys(day)
            
            fmYear = driver.find_element(
                'xpath', f'//*[@id="FromDateYear{rowNum}"]')
            fmYear.send_keys(year)
            
            toMonth = driver.find_element(
                'xpath', f'//*[@id="ToDateMonth{rowNum}"]')
            toMonth.send_keys(month)
            
            toDay = driver.find_element(
                'xpath', f'//*[@id="ToDateDay{rowNum}"]')
            toDay.send_keys(day)  
            
            toYear = driver.find_element(
                'xpath', f'//*[@id="ToDateYear{rowNum}"]')
            toYear.send_keys(year)

    if len(dates) >= 11:
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
        try:
            submitButton.click()
        except TimeoutException:
            return -1
    else:
        while driver.current_url == ub04URL:
            if stopProcess(stopFlag): return
            time.sleep(1)
    return total

def ubDownload(driver, autoDownload, memberName, stopFlag):
    if not autoDownload:
        return
    pendingURL = 'https://www.officeally.com/secure_oa.asp?GOTO=OnlineEntry&TaskAction=Pending&Msg=RCL'
    driver.get(pendingURL)

    if stopProcess(stopFlag): return

    claimID = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(('xpath', '//*[@id="tblClaims"]/tbody/tr[2]/td[4]')))
    claimID = claimID.text

    pdfURL = f'https://www.officeally.com/oa/Claims/OA_PrintClaim_UB04.aspx?ClaimID={claimID}&FormType=UB92&Virtual=Y'
    driver.get(pdfURL)

    if stopProcess(stopFlag): return
    iframe = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="div2"]/iframe'))
    )
    driver.switch_to.frame(iframe)

    openButton = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable(('xpath', '//*[@id="open-button"]'))
    )
    if stopProcess(stopFlag): return
    openButton.click()

    time.sleep(1)

    pdfList = glob.glob(os.path.join(autoDownload, '*.pdf'))
    renamedPDF = f'{memberName}.pdf'
    renamedPDF = os.path.join(autoDownload, renamedPDF)

    if renamedPDF in pdfList:
        suffix = 2
        while True:
            renamedPDF = f'{memberName} {suffix}.pdf'
            renamedPDF = os.path.join(autoDownload, renamedPDF)
            if renamedPDF not in pdfList:
                break
            suffix += 1

    else:
        renamedPDF = os.path.join(autoDownload, renamedPDF)

    curPDF = max(pdfList, key=os.path.getctime)
    os.rename(curPDF, renamedPDF)