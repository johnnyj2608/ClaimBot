from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from .claimFormsHelper import stopProcess
from .cmsForm import cmsScript
from .ubForm import ubScript
import time

def login(driver, officeAllyURL, username, password, stopFlag):
    if username and password:
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
    while driver.current_url != officeAllyURL:
        if stopProcess(stopFlag): return
        time.sleep(1)

def formatPath(path):
    if path:
        return path.replace('/', '\\')

def officeAllyAutomate(form, 
                       members, 
                       start, 
                       end, 
                       filePath,
                       autoSubmit,
                       autoDownloadPath,
                       statusLabel,
                       stopFlag,
                       updateSummary, 
                       callback):
    try:     
        officeAllyURL = 'https://www.officeally.com/secure_oa.asp'

        options = webdriver.ChromeOptions()
        options.add_experimental_option("detach", True)
        options.add_experimental_option('prefs', {
            "download.default_directory": formatPath(autoDownloadPath), 
            "download.prompt_for_download": False, 
            "download.directory_upgrade": True,
            "plugins.always_open_pdf_externally": True,
            })

        driver = webdriver.Chrome(options=options)
        driver.get(officeAllyURL)
        driver.maximize_window()

        login(driver, officeAllyURL, form['username'], form['password'], stopFlag)

        if form['form'] == "Professional (CMS)":
            cmsScript(driver,
                      form,
                      members,
                      start, 
                      end,
                      filePath,
                      autoSubmit,
                      autoDownloadPath,
                      statusLabel,
                      updateSummary,
                      stopFlag)
        else:
            ubScript(driver,
                      form,
                      members,
                      start, 
                      end,
                      filePath,
                      autoSubmit,
                      autoDownloadPath,
                      statusLabel,
                      updateSummary,
                      stopFlag)

    except Exception as e:
        print("An error occurred:", str(e))
        driver.quit()
        statusLabel.configure(text=f"Error has occurred", text_color="red")
        statusLabel.update()
    finally:
        callback()
        pendingURL = 'https://www.officeally.com/secure_oa.asp?GOTO=OnlineEntry&TaskAction=Pending&Msg=RCL'
        driver.get(pendingURL)