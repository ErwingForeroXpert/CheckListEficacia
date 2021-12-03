from logging import warning
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
import glob

def waitElement(driver, element, by=By.ID, exist=False):
    return WebDriverWait(driver, 30).until(
        lambda driver: driver.find_element(by,element) if exist else EC.visibility_of_element_located((by, element))
    )


def waitElementDisable(driver, element, by=By.ID):
    return WebDriverWait(driver, 30).until(
        EC.invisibility_of_element_located((by, element))
    )

def elementIsVisible(driver, element, by=By.ID):
    try:
        return driver.find_element(by, element).is_displayed()
    except Exception as e:
        warning(e)
        return False
        

def waitDownload(path):
    if not os.path.exists(path):
        os.makedirs(path)

    tempfiles = 0
    while tempfiles == 0:
        time.sleep(1)
        for fname in os.listdir(path):
            if "crdownload" in fname or "tmp" in fname:
                tempfiles = 0  
                break
            else:
                tempfiles = 1


def deleteTemporals(path):
    if not os.path.exists(path):
        return
    for fname in os.listdir(path):
        os.remove(os.path.join(path, fname))


def getMostRecentFile(path, _filter=None):
    list_of_files = glob.glob(fr'{path}/*')
    list_of_files = _filter(list_of_files) if _filter is not None else list_of_files
    return max(list_of_files, key=os.path.getctime)
