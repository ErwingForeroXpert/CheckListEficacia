from utils import waitDownload, waitElement, waitElementDisable, deleteTemporals, elementIsVisible, getMostRecentFile, \
    insertInLog, foundInErrorMessages, clickAlert
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select

from datetime import datetime, timedelta
from dotenv import dotenv_values
from selenium import webdriver
import xlwings as xw
from logging import exception
import time
import os
import pymsgbox


config = dotenv_values("./Program/.env")
files_route = fr"{os.path.dirname(os.path.realpath(__file__))}\files"

# insertar la direccion de descarga
chromeOptions = webdriver.ChromeOptions()
prefs = {"download.default_directory": files_route}
chromeOptions.add_experimental_option("prefs", prefs)
chromeOptions.add_argument("start-maximized")


def search(driver, text):
    """Search in Eficacia platform

    Args:
        driver ([type]): Chrome driver manager
        text ([String]): information that wil be search
    """
    search_txt_element = driver.find_element_by_xpath(
        "//div[@id='searchBox']/b/input[1]")
    search_txt_element.click()
    search_txt_element.send_keys(text)

    search_btn_element = driver.find_element_by_xpath(
        "//div[@id='submitSearch']/input[@type='image']")
    search_btn_element.click()


def login(driver):
    user_element = driver.find_element_by_xpath("//input[@id='username']")
    user_element.send_keys(config["USER"])

    pass_element = driver.find_element_by_xpath("//input[@id='password']")
    pass_element.send_keys(config["PASSWORD"])
    pass_element.send_keys(Keys.ENTER)


def runMacro(nameMacro, _args=None):
    # ejecutar macro
    result = None
    parent_folder = os.path.join(os.getcwd(), "Files")
    _book = None
    for fname in os.listdir(parent_folder):
        if "xlsm" in fname and "~" not in fname:
            _book = xw.Book(os.path.join(parent_folder, fname))
            break

    if _book is None:
        raise Exception(
            f"en la ruta {parent_folder}, el libro no se encuentra o no se puede abrir")

    result = _book.macro(nameMacro)(
        *_args) if _args is not None else _book.macro(nameMacro)()

    _book.save()

    if len(_book.app.books) == 1:
        _book.app.quit()
    else:
        _book.close()

    return result


def returnHomeFrame(driver, switch_default=False):
    if switch_default:
        driver.switch_to.default_content()
    waitElement(
        driver,
        "//frame[@name='mainFrame']",
        By.XPATH)
    driver.switch_to.frame("mainFrame")


def click_option(_select, option):
    options_type_document = Select(_select).options
    for _op in options_type_document:
        if _op.text == option:
            _op.click()


def validateErrorMessage(element, extra_info=""):
    if foundInErrorMessages(element.text):
        insertInLog(f"Error encontrado {element.text} {extra_info}")
        return True
    else:
        return False


def downloadIncomeFile(driver):
    waitElement(driver, "//table[@class='tablaexhibir']", By.XPATH)
    driver.find_element_by_xpath(
        "//table[@class='tablaexhibir TablaContainerTable']/tbody/tr/td[@class='td2']//a[contains(text(), 'Detalles_Documentos')]").click()

    returnHomeFrame(driver, True)
    driver.switch_to.frame("central")
    waitElement(driver,
                "//select[contains(@name,'tipo_documento')]", By.XPATH)
    click_option(driver.find_element(
        By.XPATH, "//select[contains(@name,'tipo_documento')]"), "Series-Recepciones")

    _today = datetime.now()
    _date_before = _today - timedelta(days=25)
    driver.find_element_by_name("fecha_final").clear()
    driver.find_element_by_name(
        "fecha_final").send_keys(_today.strftime('%m/%d/%Y'))
    driver.find_element_by_name("fecha_inicial").clear()
    driver.find_element_by_name("fecha_inicial").send_keys(
        _date_before.strftime('%m/%d/%Y'))
    driver.find_element(
        By.XPATH, "//input[contains(@type,'submit') and contains(@value,'Continuar')]").click()
    clickAlert(driver)

    returnHomeFrame(driver, True)
    driver.switch_to.frame("central")
    waitElementDisable(
        driver, "//input[contains(@type,'submit') and contains(@value,'Continuar')]", By.XPATH)

    if elementIsVisible(chrome_driver, "//a/small[contains(text(),'Descargar Archivo')]", By.XPATH):
        driver.find_element(
            By.XPATH, "//a/small[contains(text(),'Descargar Archivo XLS')]").click()
        waitDownload(files_route)
        return True
    else:
        _message = f"No se entro archivo de ingresos, fecha: {_today.strftime('%m/%d/%Y')} - {_date_before.strftime('%m/%d/%Y')}"
        validateErrorMessage(driver.find_element(By.XPATH, "//div[contains(@class,'mensaje')]"), _message)
        pymsgbox.alert(_message)
        return False


def downloadBalanceFile(driver):
    waitElement(driver, "//table[@class='tablaexhibir']", By.XPATH)
    driver.find_element_by_xpath(
        "//table[@class='tablaexhibir TablaContainerTable']/tbody/tr/td[@class='td2']//a[contains(text(), 'detalles saldos')]").click()

    returnHomeFrame(driver, True)
    driver.switch_to.frame("central")
    waitElement(driver,
                "//select[contains(@name,'tipo_producto')]", By.XPATH)

    driver.find_element(
        By.XPATH, "//input[contains(@type,'submit') and contains(@value,'Continuar')]").click()

    returnHomeFrame(driver, True)
    driver.switch_to.frame("central")
    
    waitElementDisable(
        driver, "//input[contains(@type,'submit') and contains(@value,'Continuar')]", By.XPATH)

    if elementIsVisible(chrome_driver, "//a/small[contains(text(),'Descargar Archivo')]", By.XPATH):
        driver.find_element(
            By.XPATH, "//a/small[contains(text(),'Descargar Archivo XLS')]").click()
        waitDownload(files_route)
        return True
    else:
        _message = f"No se entro archivo de saldos"
        validateErrorMessage(driver.find_element(By.XPATH, "//div[contains(@class,'mensaje')]"), _message)
        pymsgbox.alert(_message)
        return False


if __name__ == "__main__":
    chrome_driver = webdriver.Chrome(
        executable_path=ChromeDriverManager().install(), options=chromeOptions)
    try:
        deleteTemporals(files_route)
        chrome_driver.get(config["URL_EFICACIA"])

        # Esperar Login
        returnHomeFrame(chrome_driver)

        waitElement(chrome_driver, "//input[@id='username']", By.XPATH, True)
        login(chrome_driver)

        if elementIsVisible(chrome_driver, "//div[@class='error login_error' and contains(text(), 'YA TIENE SESION ACTIVA')]", By.XPATH):
            close_sessions_element = chrome_driver.find_element_by_xpath(
                "//input[@type='checkbox'][@name='cerrar_sesiones_anteriores']")
            chrome_driver.execute_script(
                "arguments[0].setAttribute('checked',arguments[1])", close_sessions_element, True)
            login(chrome_driver)

        # Esperar el Home
        waitElement(
            chrome_driver,
            "//div[@id='Layer1']",
            By.XPATH)

        # Buscar Ficha tecnica
        waitElement(
            chrome_driver,
            "//frame[@name='izquierda']",
            By.XPATH)
        chrome_driver.switch_to.frame("izquierda")
        search(chrome_driver, "ficha")

        returnHomeFrame(chrome_driver, True)

        # buscar en el frame central
        waitElement(
            chrome_driver,
            "//frame[@name='central']",
            By.XPATH)
        chrome_driver.switch_to.frame("central")

        waitElement(chrome_driver, "//table[@class='tablaexhibir']", By.XPATH)
        chrome_driver.find_element_by_xpath(
            "//table[@class='tablaexhibir TablaContainerTable']/tbody/tr/td[@class='td2'][@align='left']/font").click()

        # consultar los articulos
        waitElement(chrome_driver,
                    "//table/b[contains='*Son Campos Obligatorios']")
        chrome_driver.find_element_by_xpath(
            "//table//input[@type='submit']").click()

        # descargar el archivo
        waitElementDisable(
            chrome_driver, "//table//input[@type='submit']", By.XPATH)
        chrome_driver.find_element_by_xpath(
            "//a/small[contains(text(),'Archivo XLS')]").click()
        time.sleep(1)
        waitDownload(files_route)

        initiatives_file = getMostRecentFile(
            files_route, lambda x: [v for v in x if "xls" in v.lower()])

        # Buscar Detalles documentos
        returnHomeFrame(chrome_driver, True)
        chrome_driver.switch_to.frame("izquierda")
        search(chrome_driver, "Detalles Documentos")

        # buscar en el frame central
        returnHomeFrame(chrome_driver, True)
        chrome_driver.switch_to.frame("central")

        # descargar archivo ingresos
        _result = downloadIncomeFile(chrome_driver)
        if _result:
            income_file = getMostRecentFile(
                files_route, lambda x: [v for v in x if "xls" in v.lower()])

        # Buscar Detalles documentos
        returnHomeFrame(chrome_driver, True)
        chrome_driver.switch_to.frame("izquierda")
        search(chrome_driver, "Detalles Documentos")

        # buscar en el frame central
        returnHomeFrame(chrome_driver, True)
        chrome_driver.switch_to.frame("central")

        # descargar archivo saldos
        _result = downloadBalanceFile(chrome_driver)
        if _result:
            balance_file = getMostRecentFile(
                files_route, lambda x: [v for v in x if "xls" in v.lower()])

        chrome_driver.close()
        

        path_init = '\\'.join(initiatives_file.split('\\')[:-1])
        # necesario para las formulas de excel
        path_end = '\\' + "[" + initiatives_file.split('\\')[-1] + "]"
        initiatives_file = fr"{path_init}{path_end}"

        # ejecutar macro eliminar iniciativas completas
        runMacro('Módulo1.EliminarIniciativasCompletas')
        # ejecutar macro para actualizar iniciativas
        runMacro('Módulo1.ActualizarIniciativas', [initiatives_file])

        pymsgbox.alert("\n Proceso Terminado, ya puede cerrar la ventana \n")
        print("\n Proceso Terminado, ya puede cerrar la ventana \n")

    except Exception as e:
        exception(e)
