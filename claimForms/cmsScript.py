from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
from scheduleParser import getDatesFromWeekdays

def cmsScript(driver, 
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
        memberName = lastName+', '+firstName+' ['+birthDate.strftime("%m/%d/%Y")+']'
        
        if cmsStored(driver, insurance, summary, memberName):
            dates = getDatesFromWeekdays(start, end, schedule, authStart, authEnd)
            cmsForm(driver, dates, autoSubmit, stopFlag)

        completedMembers += 1
        statusLabel.configure(text=f"Completed Members: {completedMembers}/{totalMembers}")
        statusLabel.update()

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
    
    renderingProviderCombo.select_by_visible_text(summary['renderingProvider'])
    facilitiesCombo.select_by_visible_text(summary['facilities'])
    templatesCombo.select_by_visible_text(insurance)

    createClaimButton.click()
    return True

def cmsForm(driver, dates, autoSubmit, stopFlag):
    driver.switch_to.default_content()
    iframe = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="Iframe9"]'))
    )
    driver.switch_to.frame(iframe)

    addRowButton = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="btnAddRow"]'))
    )

    placeDefault = driver.find_element(
            'xpath', 
            '//*[@id="ctl00_phFolderContent_ucHCFA_ucHCFALineItem_ucClaimLineItem_PLACE_OF_SVC0"]')
    placeDefault = placeDefault.get_attribute("value")

    cptDefault = driver.find_element(
            'xpath', 
            '//*[@id="ctl00_phFolderContent_ucHCFA_ucHCFALineItem_ucClaimLineItem_CPT_CODE0"]')
    cptDefault = cptDefault.get_attribute("value")

    chargeDefault = driver.find_element(
            'xpath', 
            '//*[@id="ctl00_phFolderContent_ucHCFA_ucHCFALineItem_ucClaimLineItem_DOS_CHRG0"]')
    chargeDefault = chargeDefault.get_attribute("value")

    unitsDefault = driver.find_element(
            'xpath', 
            '//*[@id="ctl00_phFolderContent_ucHCFA_ucHCFALineItem_ucClaimLineItem_UNITS0"]')
    unitsDefault = unitsDefault.get_attribute("value")

    for rowNum in range(12, len(dates)):
        addRowButton.click()

        placeRow = driver.find_element(
            'xpath', 
            f'//*[@id="ctl00_phFolderContent_ucHCFA_ucHCFALineItem_ucClaimLineItem_PLACE_OF_SVC{rowNum}"]'
            )
        
        cptRow = driver.find_element(
            'xpath', 
            f'//*[@id="ctl00_phFolderContent_ucHCFA_ucHCFALineItem_ucClaimLineItem_CPT_CODE{rowNum}"]'
            )
        
        chargeRow = driver.find_element(
            'xpath', 
            f'//*[@id="ctl00_phFolderContent_ucHCFA_ucHCFALineItem_ucClaimLineItem_DOS_CHRG{rowNum}"]'
            )
        
        unitsRow = driver.find_element(
            'xpath', 
            f'//*[@id="ctl00_phFolderContent_ucHCFA_ucHCFALineItem_ucClaimLineItem_UNITS{rowNum}"]'
            )
        
        placeRow.send_keys(placeDefault)
        cptRow.send_keys(cptDefault)
        chargeRow.send_keys(chargeDefault)
        unitsRow.send_keys(unitsDefault)

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

    if stopFlag.value:
        stopFlag.value = False
    elif autoSubmit:
        pass
    else:
        pass
        # Wait for human input