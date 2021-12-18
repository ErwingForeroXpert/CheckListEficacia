from datetime import datetime
from logging import warning
import xlwings as xw
import logging
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
import glob
import re

def waitElement(driver, element, by=By.ID, exist=False):
    """wait for the element appear on the Screen

    Args:
        driver (WebDriver): web driver see https://www.selenium.dev/documentation/webdriver/
        element (Object<Any>): element see https://selenium-python.readthedocs.io/locating-elements.html
        by (String, optional): Searcher. Defaults to By.ID.
        exist (bool, optional): if element already existed. Defaults to False.

    Returns:
        WebDriverWait: Constructor, takes a WebDriver instance and timeout in seconds.
    """

    return WebDriverWait(driver, 30).until(
        lambda driver: driver.find_element(by,element) if exist else EC.visibility_of_element_located((by, element))
    )


def waitElementDisable(driver, element, by=By.ID):
    """wait for the element disappear of the screen

    Args:
        driver (WebDriver): web driver see https://www.selenium.dev/documentation/webdriver/
        element (Object<Any>): element see https://selenium-python.readthedocs.io/locating-elements.html
        by (String, optional): Searcher. Defaults to By.ID.

    Returns:
        WebDriverWait: Constructor, takes a WebDriver instance and timeout in seconds.
    """

    return WebDriverWait(driver, 30).until(
        EC.invisibility_of_element_located((by, element))
    )

def elementIsVisible(driver, element, by=By.ID, wait=2):
    """Validate if element is visible

    Args:
        driver (WebDriver): web driver see https://www.selenium.dev/documentation/webdriver/
        element (Object<Any>): element see https://selenium-python.readthedocs.io/locating-elements.html
        by (String, optional): Searcher. Defaults to By.ID.
        wait (int, optional): delay before starting the process. Defaults to 2.

    Returns:
        bool: if element is visible
    """
    time.sleep(wait)
    try:
        return driver.find_element(by, element).is_displayed()
    except Exception as e:
        warning(e)
        return False
        
def waitBookDisable(mybook):
    """Wait until the book is closed

    Args:
        mybook (String): path of the excel book
    """
    with xw.App() as app:
        while len(app.books) != 0:
            time.sleep(1)
            for book in app.books:
                if book.fullname == mybook:
                    book.close()
            
def waitDownload(path):
    """wait that elements in state downloading disappear

    Args:
        path (String): Folder of downloads
    """
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
    """Create folders

    Args:
        path (String): Folder parents
        folders (String): Folder childrens
    """
    for folder in folders:
        if not os.path.exists(os.path.join(path, folder)):
            os.makedirs(os.path.join(path, folder))

def deleteTemporals(path):
    """Delete elements of path

    Args:
        path (String): Folder parent
    """
    for fname in os.listdir(path):
        os.remove(os.path.join(path, fname))


def getMostRecentFile(path, _filter=None):
    """Get most recent file for date

    Args:
        path (String): Folder parent
        _filter (function, optional): filter name of documents. Defaults to None.

    Returns:
        String: path of file most recent
    """
    list_of_files = glob.glob(fr'{path}/*')
    list_of_files = _filter(list_of_files) if _filter is not None else list_of_files
    return max(list_of_files, key=os.path.getctime)

def foundInErrorMessages(message):
    """validate if the message is found in the error messages

    Args:
        message (String): message to find

    Returns:
        bool: found in the error messages
    """
    return message in [
        "2001: NO EXISTEN REGISTROS EN LA CONSULTA"
    ]

def clickAlert(driver):
    """Click alert message if appear

    Args:
        driver (WebDriver): web driver see https://www.selenium.dev/documentation/webdriver/
    """
    try:
        WebDriverWait(driver, 5).until (EC.alert_is_present())
        alert = driver.switch_to.alert
        alert.accept()
    except Exception as e:
        insertInLog("alert does not Exist in page", "info")

def insertInLog(message, type="debug"):
    """Insert new line in file log

    Args:
        message (String): message
        type (str, optional): type of log. Defaults to "debug".
    """
    logging.basicConfig(filename='checklisteficacia.log', encoding='utf-8', filemode='w', format='%(name)s - %(levelname)s - %(message)s')
    loger = {
        "debug": logging.debug,
        "warning": logging.warning,
        "info": logging.info,
        "error": logging.error,
    }[type]

    loger(f"{datetime.now().strftime(r'%d/%m/%Y %H:%M:%S')} {message} \n")
    print(f"{datetime.now().strftime(r'%d/%m/%Y %H:%M:%S')} {message} \n")

def ifErrorFalse(cb, *args):
    """if generate error return false

    Args:
        cb (function): function

    Returns:
        Any: function return or False
    """
    try:
        return cb(*args)
    except Exception as e:
        return False

def isEmpty(value):
    """if is empty value

    Args:
        value (String|int): posible empty value 

    Returns:
        bool: if is empty value
    """
    try:
        return str(value).strip() == ""
    except Exception as e:
        return True

def string2Number(value):
    """Convert String to Number

    Args:
        value (String): [description]

    Raises:
        Exception: [description]

    Returns:
        [type]: [description]
    """
    value = value if isinstance(value, str) else str(value)
    _temp_val = re.findall(r"\d+", value) 
    if _temp_val is not []:
        return int(_temp_val[0])
    else:
        raise Exception(f"{value} has not number")

def exceptionHandler(func):
    def inner_function(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            _message = f"{func.__name__} - {e}"
            insertInLog(_message, "error")
            raise Exception(_message)
    return inner_function