import datetime
from logging import warning
import logging
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
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

def elementIsVisible(driver, element, by=By.ID, wait=2):
    time.sleep(wait)
    try:
        return driver.find_element(by, element).is_displayed()
    except Exception as e:
        warning(e)
        return False
        

def waitDownload(path):
    tempfiles = 0
    while tempfiles == 0:
        time.sleep(1)
        for fname in os.listdir(path):
            if "crdownload" in fname or "tmp" in fname:
                tempfiles = 0  
                break
            else:
                tempfiles = 1

def createNecesaryFolders(path, folders):
    for folder in folders:
        if not os.path.exists(os.path.join(path, folder)):
            os.makedirs(os.path.join(path, folder))

def deleteTemporals(path):
    for fname in os.listdir(path):
        os.remove(os.path.join(path, fname))


def getMostRecentFile(path, _filter=None):
    list_of_files = glob.glob(fr'{path}/*')
    list_of_files = _filter(list_of_files) if _filter is not None else list_of_files
    return max(list_of_files, key=os.path.getctime)

def foundInErrorMessages(message):
    return message in [
        "2001: NO EXISTEN REGISTROS EN LA CONSULTA"
    ]

def clickAlert(driver):
    try:
        WebDriverWait(driver, 5).until (EC.alert_is_present())
        alert = driver.switch_to.alert
        alert.accept()
    except TimeoutException:
        insertInLog("alert does not Exist in page", "info")

def insertInLog(message, type="debug"):
    logging.basicConfig(filename='checklisteficacia.log', encoding='utf-8', filemode='w', format='%(name)s - %(levelname)s - %(message)s')
    loger = {
        "debug": logging.debug,
        "warning": logging.warning,
        "info": logging.info,
        "error": logging.error,
    }[type]

    loger(f"{datetime.now().strftime(r'%d/%m/%Y %H:%M:%S')} {message} \n")