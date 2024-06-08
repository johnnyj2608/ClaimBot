from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from .cmsForm import cmsScript
from .ubForm import ubScript

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
                       callback):
    try:     
        office_ally = 'https://www.officeally.com/secure_oa.asp'

        options = webdriver.ChromeOptions()
        options.add_experimental_option("detach", True)
        options.add_experimental_option('prefs', {
            "download.default_directory": formatPath(autoDownloadPath), 
            "download.prompt_for_download": False, 
            "download.directory_upgrade": True,
            "plugins.always_open_pdf_externally": True,
            })

        driver = webdriver.Chrome(options=options)
        driver.get(office_ally)
        driver.maximize_window()

        login(driver, form['username'], form['password'])

        submissionSummary = None
        if form['form'] == "Professional (CMS)":
            submissionSummary = cmsScript(driver,
                      form,
                      members,
                      start, 
                      end,
                      filePath,
                      autoSubmit,
                      autoDownloadPath,
                      statusLabel,
                      stopFlag)
        else:
            submissionSummary = ubScript(driver,
                      form,
                      members,
                      start, 
                      end,
                      filePath,
                      autoSubmit,
                      autoDownloadPath,
                      statusLabel,
                      stopFlag)

        statusLabel.configure(text_color="green")
        statusLabel.update()

    except Exception as e:
        print("An error occurred:", str(e))
        driver.quit()
        statusLabel.configure(text=f"Error has occurred", text_color="red")
        statusLabel.update()
    finally:
        callback(submissionSummary)
        pendingURL = 'https://www.officeally.com/secure_oa.asp?GOTO=OnlineEntry&TaskAction=Pending&Msg=RCL'
        driver.get(pendingURL)