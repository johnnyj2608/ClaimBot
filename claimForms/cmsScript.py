from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
from claimForms.claimFormsHelper import *
from excel import recordClaims
import time
import os
import glob

def cmsScript(driver, 
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
            memberSearch = memberName+' ['+member['birthDate'].strftime("%m/%d/%Y")+']'
            
            total = -1
            if cmsStored(driver, summary, memberSearch):
                dates = getDatesFromWeekdays(start, end, member['schedule'], member['authStart'], member['authEnd'])
                dates = intersectVacations(dates, member['vacationStart'], member['vacationEnd'])
                total = cmsForm(driver, summary, member['authID'], member['dxCode'], 
                                start, end, dates, autoSubmit, stopFlag)
            if stopProcess(stopFlag): return
            if total == -1:
                summaryStats['failed'] += 1
            else:
                summaryStats['submitted'] += 1
                cmsDownload(driver, autoDownload, memberName, stopFlag)
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

def cmsStored(driver, summary, memberName):
    storedInfoURL='https://www.officeally.com/secure_oa.asp?GOTO=OnlineEntry&TaskAction=Manage'
    if driver.current_url != storedInfoURL:
        driver.get(storedInfoURL)
        iframe = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(('xpath', '//*[@id="Iframe9"]'))
        )
        driver.switch_to.frame(iframe)

    try:
        patientsCombo = Select(WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ddlPatient"]'))
    ))
        patientsCombo.select_by_visible_text(memberName)
    except NoSuchElementException:
        print(f"Element with visible text '{memberName}' not found.")
        return False
    
    try:
        payerCombo = Select(WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(('xpath', '//*[@id="ddlPayer"]'))))
        payerCombo.select_by_visible_text(summary['payer'])

        billingProviderCombo = Select(WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(('xpath', '//*[@id="ddlBillingProvider"]'))))
        billingProviderCombo.select_by_visible_text(summary['billingProvider'])
        
        renderingProviderCombo = Select(WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(('xpath', '//*[@id="ddlRenderingProvider"]'))))
        renderingProviderCombo.select_by_visible_text(summary['renderingProvider'])

        facilitiesCombo = Select(WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(('xpath', '//*[@id="ddlFacility"]'))))
        facilitiesCombo.select_by_visible_text(summary['facilities'])

    except NoSuchElementException:
        print("Failed to select stored information")
        return False
    
    createClaimButton = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="Button2"]')))
    createClaimButton.click()
    return True

def cmsForm(driver, summary, authID, dxCode, start, end, dates, autoSubmit, stopFlag):
    if stopProcess(stopFlag): return
    cms1500URL = driver.current_url
    driver.switch_to.default_content()
    iframe = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="Iframe9"]'))
    )
    driver.switch_to.frame(iframe)

    outsideLab = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucHCFA_OUTSIDE_LAB2"]')))
    outsideLab.click()

    authField = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucHCFA_PRIOR_AUTH_NUMBER"]')))
    authField.send_keys(authID)

    dxField = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucHCFA_DIAGNOSIS_CODECMS0212_1"]')))
    dxField.send_keys(dxCode)

    acceptAssignment = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucHCFA_ACCEPT_SIGNATURE1"]')))
    acceptAssignment.click()

    statementFromMonth = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucHCFA_DateInitialTreatment_Month"]')))
    statementFromMonth.send_keys(start.month)

    statementFromDay = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucHCFA_DateInitialTreatment_Day"]')))
    statementFromDay.send_keys(start.day)
    
    statementFromYear = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucHCFA_DateInitialTreatment_Year"]')))
    statementFromYear.send_keys(start.year)
    
    statementToMonth = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucHCFA_DateLastSeen_Month"]')))
    statementToMonth.send_keys(end.month)

    statementToDay = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucHCFA_DateLastSeen_Day"]')))
    statementToDay.send_keys(end.day)

    statementToYear = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucHCFA_DateLastSeen_Year"]')))
    statementToYear.send_keys(end.year)

    copyPatient = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="lnkPatientCopy"]')))
    copyPatient.click()

    confirmPatient = WebDriverWait(driver, 100).until(
        EC.element_to_be_clickable(('xpath', '/html/body/div[5]/div[3]/div/button[1]')))
    confirmPatient.click()

    addRowButton = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="btnAddRow"]')))

    for rowNum in range(len(dates)):
        if stopProcess(stopFlag): return
        if rowNum > 11:
            addRowButton.click()

        placeRow = driver.find_element(
            'xpath', 
            f'//*[@id="ctl00_phFolderContent_ucHCFA_ucHCFALineItem_ucClaimLineItem_PLACE_OF_SVC{rowNum}"]')
        placeRow.send_keys(summary['servicePlace'])
        
        cptRow = driver.find_element(
            'xpath', 
            f'//*[@id="ctl00_phFolderContent_ucHCFA_ucHCFALineItem_ucClaimLineItem_CPT_CODE{rowNum}"]')
        cptRow.send_keys(summary['cptCode'])
        
        modifierRow = driver.find_element(
            'xpath', 
            f'//*[@id="ctl00_phFolderContent_ucHCFA_ucHCFALineItem_ucClaimLineItem_MODIFIER_A{rowNum}"]')
        modifierRow.send_keys(summary['modifier'])
        
        diagnosisRow = driver.find_element(
            'xpath', 
            f'//*[@id="ctl00_phFolderContent_ucHCFA_ucHCFALineItem_ucClaimLineItem_DOS_DIAG_CODE{rowNum}"]')
        diagnosisRow.send_keys(summary['diagnosis'])
        
        chargeRow = driver.find_element(
            'xpath', 
            f'//*[@id="ctl00_phFolderContent_ucHCFA_ucHCFALineItem_ucClaimLineItem_DOS_CHRG{rowNum}"]')
        chargeRow.send_keys(summary['charges'])
        
        unitsRow = driver.find_element(
            'xpath', 
            f'//*[@id="ctl00_phFolderContent_ucHCFA_ucHCFALineItem_ucClaimLineItem_UNITS{rowNum}"]')
        unitsRow.send_keys(summary['units'])
        
    for rowNum in range(len(dates)):
        if stopProcess(stopFlag): return
        curDate = dates[rowNum]
        month = curDate.month
        day = curDate.day
        year = curDate.year

        fmMonth = driver.find_element(
            'xpath', 
            f'//*[@id="ctl00_phFolderContent_ucHCFA_ucHCFALineItem_ucClaimLineItem_FM_DATE_OF_SVC_MONTH{rowNum}"]')
        fmMonth.send_keys(month)
        
        fmDay = driver.find_element(
            'xpath', 
            f'//*[@id="ctl00_phFolderContent_ucHCFA_ucHCFALineItem_ucClaimLineItem_FM_DATE_OF_SVC_DAY{rowNum}"]')
        fmDay.send_keys(day)
        
        fmYear = driver.find_element(
            'xpath', 
            f'//*[@id="ctl00_phFolderContent_ucHCFA_ucHCFALineItem_ucClaimLineItem_FM_DATE_OF_SVC_YEAR{rowNum}"]')
        fmYear.send_keys(year)
        
        toMonth = driver.find_element(
            'xpath', 
            f'//*[@id="ctl00_phFolderContent_ucHCFA_ucHCFALineItem_ucClaimLineItem_TO_DATE_OF_SVC_MONTH{rowNum}"]')
        toMonth.send_keys(month)
        
        toDay = driver.find_element(
            'xpath', 
            f'//*[@id="ctl00_phFolderContent_ucHCFA_ucHCFALineItem_ucClaimLineItem_TO_DATE_OF_SVC_DAY{rowNum}"]')
        toDay.send_keys(day)
        
        toYear = driver.find_element(
            'xpath', 
            f'//*[@id="ctl00_phFolderContent_ucHCFA_ucHCFALineItem_ucClaimLineItem_TO_DATE_OF_SVC_YEAR{rowNum}"]')
        toYear.send_keys(year)   

    totalField = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucHCFA_TOTAL_CHARGE"]')))
    total = totalField.get_attribute("value")

    if stopProcess(stopFlag): return
    if autoSubmit:
        submitButton = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucHCFA_btnSCUpdate"]')))
        try:
            submitButton.click()
        except TimeoutException:
            return -1
    else:
        while driver.current_url == cms1500URL:
            if stopProcess(stopFlag): return
            time.sleep(1)
    return total

def cmsDownload(driver, autoDownload, memberName, stopFlag):
    if not autoDownload:
        return
    pendingURL = 'https://www.officeally.com/secure_oa.asp?GOTO=OnlineEntry&TaskAction=Pending&Msg=RCL'
    driver.get(pendingURL)

    if stopProcess(stopFlag): return

    claimID = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(('xpath', '//*[@id="tblClaims"]/tbody/tr[2]/td[4]')))
    claimID = claimID.text

    pdfURL = f'https://www.officeally.com/IPAReport/HCFA_PrintClaim.asp?ClaimID={claimID}&FormType=HCFA&Virtual=Y&Box33=B&CMS=0805'
    driver.get(pdfURL)

    if stopProcess(stopFlag): return

    iframe = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '/html/body/iframe'))
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