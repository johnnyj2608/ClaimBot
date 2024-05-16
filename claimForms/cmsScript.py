from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
from claimForms.claimFormsHelper import *
import time

def cmsScript(driver, 
              summary, 
              members, 
              start, 
              end, 
              autoSubmit, 
              statusLabel, 
              stopFlag):
    
    totalMembers, completedMembers = len(members), 0
    submittedClaims = []
    for member in members:
        if stopProcess(stopFlag): return

        print(member)

        lastName, firstName, birthDate, medicaid, authID, dxCode, schedule, authStart, authEnd, vacaStart, vacaEnd, Exclude = member
        memberName = lastName+', '+firstName+' ['+birthDate.strftime("%m/%d/%Y")+']'
        
        total = -1
        if cmsStored(driver, summary['insurance'], summary, memberName):
            dates = getDatesFromWeekdays(start, end, schedule, authStart, authEnd)
            dates = intersectVacations(dates, start, end)
            total = cmsForm(driver, dxCode, authID, dates, autoSubmit, stopFlag)

        submittedClaims.append([lastName, firstName, total])
        completedMembers += 1
        statusLabel.configure(text=f"Completed Members: {completedMembers}/{totalMembers}")
        statusLabel.update()
    return submittedClaims

def cmsStored(driver, insurance, summary, memberName):
    storedInfoURL='https://www.officeally.com/secure_oa.asp?GOTO=OnlineEntry&TaskAction=Manage'
    if driver.current_url != storedInfoURL:
        driver.get(storedInfoURL)
        iframe = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(('xpath', '//*[@id="Iframe9"]'))
        )
        driver.switch_to.frame(iframe)

    patientsCombo = Select(WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ddlPatient"]'))
    ))
    
    try:
        patientsCombo.select_by_visible_text(memberName)
    except NoSuchElementException:
        print(f"Element with visible text '{memberName}' not found.")
        return False
    
    payerCombo = Select(WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ddlPayer"]'))
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
    
    # renderingProviderCombo.select_by_visible_text(summary['renderingProvider'])
    facilitiesCombo.select_by_visible_text(summary['facilities'])
    templatesCombo.select_by_visible_text(insurance)

    createClaimButton.click()
    return True

def cmsForm(driver, dxCode, authID, dates, autoSubmit, stopFlag):
    if stopProcess(stopFlag): return
    cms1500URL = driver.current_url
    driver.switch_to.default_content()
    iframe = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="Iframe9"]'))
    )
    driver.switch_to.frame(iframe)

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

    # If more than 12 dates
    for rowNum in range(12, len(dates)):
        if stopProcess(stopFlag): return
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
        
    # If less than 12 dates
    for rowNum in range(11, len(dates)-1, -1):
        if stopProcess(stopFlag): return
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
        if stopProcess(stopFlag): return
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

    if stopProcess(stopFlag): return
    if autoSubmit:
        submitButton = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(('xpath', '//*[@id="ctl00_phFolderContent_ucHCFA_btnSCUpdate"]')))
        submitButton.click()
        # Wait to see next page before returning
    else:
        while driver.current_url == cms1500URL:
            time.sleep(1)
    return total