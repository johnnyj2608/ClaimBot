from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from claimForms.cmsScript import cmsScript
from claimForms.ubScript import ubScript
from claimForms.claimFormsHelper import stopProcess
from excel import recordClaims

import traceback

def login(driver, username, password):
    usernameField = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="username"]'))
    )
    passwordField = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '//*[@id="password"]'))
    )
    loginButton = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(('xpath', '/html/body/main/section/div/div/div/form/div[2]/button'))
    )

    usernameField.send_keys(username)
    passwordField.send_keys(password)
    loginButton.click()

def officeAllyAutomate(summary, 
                       members, 
                       start, 
                       end, 
                       autoSubmit,
                       autoDownloadPath,
                       statusLabel,
                       stopFlag, 
                       callback):
    try:     
        office_ally = 'https://www.officeally.com/secure_oa.asp'

        options = webdriver.ChromeOptions()
        options.add_experimental_option("detach", True)
        driver = webdriver.Chrome(options=options)
        driver.get(office_ally)
        driver.maximize_window()

        login(driver, summary['username'], summary['password'])

        if summary['form'] == "Professional (CMS)":
            submittedClaims = cmsScript(driver,
                      summary,
                      members,
                      start, 
                      end,
                      autoSubmit,
                      statusLabel,
                      stopFlag)
        else:
            submittedClaims = ubScript(driver,
                      summary,
                      members,
                      start, 
                      end,
                      autoSubmit,
                      statusLabel,
                      stopFlag)

        if stopProcess(stopFlag): return
        statusLabel.configure(text_color="green")
        statusLabel.update()

    except Exception as e:
        print("An error occurred:", str(e))
        traceback.print_exc()
        driver.quit()
        statusLabel.configure(text=f"Error has occurred", text_color="red")
        statusLabel.update()
    finally:
        callback()