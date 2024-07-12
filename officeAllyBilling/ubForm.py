from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
from .claimFormsHelper import *
from excel import recordClaims
import time
import glob
import os

def ubScript(driver, 
              form, 
              members, 
              start, 
              end, 
              filePath,
              autoSubmit,
              autoDownload,
              statusLabel,
              updateSummary,
              stopFlag):
    
    totalMembers, completedMembers = len(members), 0
    
    summary = {
        'members': 0,
        'success': 0,
        'total': 0,
        'unsubmitted': [],
    }

    for member in members:
        if stopProcess(stopFlag): return

        dates = getDatesFromWeekdays(start, end, member['schedule'], member['authStart'], member['authEnd'], 
                                             member['vacationStart'], member['vacationEnd'])
        memberName = member['lastName']+', '+member['firstName']

        summary['members'] += 1
        statusLabel.configure(text=f"{member['id']}. {memberName} ({summary['members']}/{len(members)})")
        statusLabel.update()

        if not dates:
            summary['unsubmitted'].append(member['id']+'. '+memberName + ': No available dates')
        else:
            memberSearch = member['firstName']+' '+member['lastName']
            memberSelect = member['lastName']+', '+member['firstName']+' ['+member['birthDate'].strftime("%#m/%#d/%y")+']'

            total = -1
            if ubStored(driver, form, memberSearch, memberSelect):
                total = ubForm(driver, form, member['dxCode'], dates, autoSubmit, stopFlag)
            if stopProcess(stopFlag): return
            if total != -1:
                summary['success'] += 1
                summary['total'] += total
                ubDownload(driver, autoDownload, memberName, stopFlag)
                recordClaims(filePath, 
                         memberName,
                         start.strftime("%#m/%#d/%y")+' - '+end.strftime("%#m/%#d/%y"),
                         total)
            else:
                summary['unsubmitted'].append(member['id']+'. '+memberName + ': Failed to submit')
        updateSummary(summary)

    return summary

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
    
def ubForm(driver, summary, dxCode, dates, autoSubmit, stopFlag):
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
    statementFromMonth.send_keys(dates[0].month)
    
    statementFromDay = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucUBForm_StatementFromDate_Day"]')))
    statementFromDay.send_keys(dates[0].day)
    
    statementFromYear = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucUBForm_StatementFromDate_Year"]')))
    statementFromYear.send_keys(dates[0].year)
    
    statementToMonth = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucUBForm_StatementToDate_Month"]')))
    statementToMonth.send_keys(dates[-1].month)
    
    statementToDay = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucUBForm_StatementToDate_Day"]')))
    statementToDay.send_keys(dates[-1].day)
    
    statementToYear = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucUBForm_StatementToDate_Year"]')))
    statementToYear.send_keys(dates[-1].year)

    billTypeField = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucUBForm_BillType"]')))
    billTypeField.send_keys(summary['billType'])

    admissionDateMonth = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucUBForm_DateAdmitted_Month"]')))
    admissionDateMonth.send_keys(dates[0].month)

    admissionDateDay = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucUBForm_DateAdmitted_Day"]')))
    admissionDateDay.send_keys(dates[0].day)

    admissionDateYear = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucUBForm_DateAdmitted_Year"]')))
    admissionDateYear.send_keys(dates[0].year)
    
    admissionTypeField = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucUBForm_TypeOfAdmission"]')))
    admissionTypeField.send_keys(9)

    admissionSrcField = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucUBForm_AdmissionSource"]')))
    admissionSrcField.send_keys(9)

    statusField = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucUBForm_PatientStatus"]')))
    statusField.send_keys(30)

    dxField = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucUBForm_PrimmaryDiagnosisCode"]')))
    dxField.send_keys(dxCode)
    
    for rowNum in range(1, (len(dates)*2)+1):
        if stopProcess(stopFlag): return

        curDate = dates[(rowNum-1) // 2]
        month = curDate.month
        day = curDate.day
        year = curDate.year

        revCodeRow = driver.find_element(
            'xpath', f'//*[@id="RevCode{rowNum}"]')
        revCodeRow.send_keys(summary['revenueCode'])

        descRow = driver.find_element(
            'xpath', f'//*[@id="Description{rowNum}"]')
        descRow.send_keys(summary['descriptionSDC']) if rowNum % 2 == 1 else descRow.send_keys(summary['descriptionTrans'])

        cptRow = driver.find_element(
            'xpath', f'//*[@id="Rate{rowNum}"]')
        cptRow.send_keys(summary['cptCodeSDC']) if rowNum % 2 == 1 else cptRow.send_keys(summary['cptCodeTrans'])
        
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

        unitsRow = driver.find_element(
            'xpath', f'//*[@id="Units{rowNum}"]')
        unitsRow.send_keys(summary['unitsSDC']) if rowNum % 2 == 1 else unitsRow.send_keys(summary['unitsTrans'])

        chargeRow = driver.find_element(
            'xpath', f'//*[@id="TotalCharge{rowNum}"]')
        chargeRow.send_keys(summary['chargesSDC']) if rowNum % 2 == 1 else chargeRow.send_keys(summary['chargesTrans'])

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
    return float(total)

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

    pdfName = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located(('xpath', '//*[@id="main-message"]/h1')))
    pdfName = pdfName.text

    openButton = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable(('xpath', '//*[@id="open-button"]'))
    )
    if stopProcess(stopFlag): return
    openButton.click()

    pdfList = glob.glob(os.path.join(autoDownload, '*.pdf'))
    pdfName = os.path.join(autoDownload, pdfName)
    while pdfName not in pdfList:
        if stopProcess(stopFlag): return
        time.sleep(1)
        pdfList = glob.glob(os.path.join(autoDownload, '*.pdf'))

    renamedPDF = f'{memberName}.pdf'
    renamedPDF = os.path.join(autoDownload, renamedPDF)

    suffix = 2
    while renamedPDF in pdfList:
        if stopProcess(stopFlag): return
        renamedPDF = f'{memberName} {suffix}.pdf'
        renamedPDF = os.path.join(autoDownload, renamedPDF)
        if renamedPDF not in pdfList:
            break
        suffix += 1

    os.rename(pdfName, renamedPDF)