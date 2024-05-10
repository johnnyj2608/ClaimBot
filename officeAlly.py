from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
from scheduleParser import getDatesFromWeekdays

def officeAllyAutomate(insurance, 
                       summary, 
                       members, 
                       start, 
                       end, 
                       autoSubmit,
                       statusLabel,
                       callback):
    office_ally = 'https://www.officeally.com/secure_oa.asp'

    options = webdriver.ChromeOptions()
    options.add_experimental_option("detach", True)

    driver = webdriver.Chrome(options=options)
    driver.get(office_ally)
    driver.maximize_window()
    # driver.implicitly_wait(10)

    # ------------- LOGIN PAGE ------------

    usernameField = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="username"]'))
    )
    passwordField = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="password"]'))
    )
    loginButton = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '/html/body/main/section/div/div/div/form/div[2]/button'))
    )
    
    usernameField.send_keys(summary['username'])
    passwordField.send_keys(summary['password'])
    loginButton.click()

    # ------------- SUMMARY PAGE ------------

    driver.get('https://www.officeally.com/secure_oa.asp?GOTO=OnlineEntry&TaskAction=Manage')

    iframe = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="Iframe9"]'))
    )
    driver.switch_to.frame(iframe)

    payerCombo = Select(WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ddlPayer"]'))
    ))
    patientsCombo = Select(WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="ddlPatient"]'))
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

    totalMembers, completedMembers = len(members), 0

    for member in members:
        lastName, firstName, birthDate, authID, dxCode, schedule, authStart, authEnd = member

        searchString = lastName+', '+firstName+' ['+birthDate.strftime("%m/%d/%Y")+']'
        try:
            patientsCombo.select_by_visible_text(searchString)
        except NoSuchElementException:
            print(f"Element with visible text '{searchString}' not found.")
            completedMembers += 1
            statusLabel.configure(text=f"Completed Members: {completedMembers}/{totalMembers}")
            statusLabel.update()
            continue
        dates = getDatesFromWeekdays(start, end, schedule, authStart, authEnd)

        createClaimButton.click()

        # ------------- CLAIMS PAGE ------------

        # ------------- CMS-1500 ------------

        driver.switch_to.default_content()
        iframe = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(('xpath', '//*[@id="Iframe9"]'))
        )
        driver.switch_to.frame(iframe)

        addRowButton = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(('xpath', '//*[@id="btnAddRow"]'))
        )

        # Get template values
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

        completedMembers += 1
        statusLabel.configure(text=f"Completed Members: {completedMembers}/{totalMembers}")
        statusLabel.update()
    callback()