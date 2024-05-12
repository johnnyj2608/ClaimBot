from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
from scheduleParser import getDatesFromWeekdays

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
    # https://www.officeally.com/secure_oa.asp?GOTO=UB04OnlineEntry&TaskAction=StoredInfo