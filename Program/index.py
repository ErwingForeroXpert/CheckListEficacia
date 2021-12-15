from utils import waitDownload, waitElement, waitElementDisable, deleteTemporals, elementIsVisible, getMostRecentFile, \
    insertInLog, foundInErrorMessages, clickAlert, createNecesaryFolders
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
errors_route = fr"{os.path.dirname(os.path.realpath(__file__))}\_errors"

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

    result = _book.macro(nameMacro)(*_args) if _args is not None else _book.macro(nameMacro)()

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


def click_option(_select, option, precision="same"):
    options_type_document = Select(_select).options
    for _op in options_type_document:
        validation = _op.text.lower() == option.lower() if precision == "same" else _op.text.lower() in option.lower()
        if validation:
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

    if elementIsVisible(driver, "//a/small[contains(text(),'Descargar Archivo')]", By.XPATH, wait=5):
        driver.find_element(
            By.XPATH, "//a/small[contains(text(),'Descargar Archivo XLS')]").click()
        waitDownload(files_route)
        return True
    elif elementIsVisible(driver, "//div[contains(@class,'exhibir_tabla')]", By.XPATH):
        driver.find_element(By.XPATH, "//a[contains(@href,'tipo_exportacion=XLS')]").click()
        waitDownload(files_route)
        return True
    else:
        if elementIsVisible(driver, "//div[contains(@class,'mensaje')]", By.XPATH):
            _message = f"No se entro archivo de ingresos, fecha: {_today.strftime('%m/%d/%Y')} - {_date_before.strftime('%m/%d/%Y')}"
            validateErrorMessage(driver.find_element(
                By.XPATH, "//div[contains(@class,'mensaje')]"), _message)
        else:
            _message = f"Error al intentar descargar el archivo de ingresos, fecha: {_today.strftime('%m/%d/%Y')} - {_date_before.strftime('%m/%d/%Y')}"
            driver.get_screenshot_as_file(os.path.join(
                errors_route, f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"))
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

    if elementIsVisible(driver, "//a/small[contains(text(),'Descargar Archivo')]", By.XPATH, wait=5):
        driver.find_element(
            By.XPATH, "//a/small[contains(text(),'Descargar Archivo XLS')]").click()
        waitDownload(files_route)
        return True
    elif elementIsVisible(driver, "//div[contains(@class,'exhibir_tabla')]", By.XPATH):
        driver.find_element(By.XPATH, "//a[contains(@href,'tipo_exportacion=XLS')]").click()
        waitDownload(files_route)
        return True
    else:
        if elementIsVisible(driver, "//div[contains(@class,'mensaje')]", By.XPATH):
            _message = f"No se entro archivo de saldos"
            validateErrorMessage(driver.find_element(
                By.XPATH, "//div[contains(@class,'mensaje')]"), _message)
        else:
            _message = f"Error al intentar descargar el archivo de saldos"
            driver.get_screenshot_as_file(os.path.join(
                errors_route, f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"))
        pymsgbox.alert(_message)
        return False


def validateACTS(driver, acts):
    temp_acts_asigments = []
    waitElement(driver, "//table[@class='tablaexhibir']", By.XPATH)
    driver.find_element_by_xpath(
        "//table[@class='tablaexhibir TablaContainerTable']/tbody/tr/td[@class='td2']//a[contains(text(), 'Kardex')]").click()

    for i, act in enumerate(acts):
        if "act" in act[2].lower():
            if "bgta" in act[1].lower():
                acts[i] = validateValueACT(driver, act, by="excel")
            else:
                acts[i] = validateValueACT(driver, act)
            temp_acts_asigments.append(
                [i][:5]) #first 6 registers
           
    return temp_acts_asigments


def validateValueACT(driver, act, by="table"):

    returnHomeFrame(driver, True)
    driver.switch_to.frame("central")
    click_option(waitElement(driver,
                "//select[contains(@name,'punto')]", By.XPATH), act[1], "similar")

    _today = datetime.now()
    _date_before = _today - timedelta(days=25)
    driver.find_element_by_name("articulo").clear()
    driver.find_element_by_name(act[2])
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

    if elementIsVisible(driver, "//table[contains(@class,'tablaexhibir')]", By.XPATH, wait=2) and by == "table":
        len_table = len(chrome_driver.find_elements(By.XPATH, "//table[contains(@class,'tablaexhibir')]/tbody/tr"))
        if len_table > 2:
            total_finded = chrome_driver.find_element(By.XPATH, f"//table[contains(@class,'tablaexhibir')]/tbody/tr[{len_table}]/th[16]").text
            act[5] = int(total_finded)
    elif elementIsVisible(driver, "//a/small[contains(text(),'Descargar Archivo')]", By.XPATH):
        driver.find_element(
            By.XPATH, "//a/small[contains(text(),'Descargar Archivo XLS')]").click()
        waitDownload(files_route)
        temp_file_act = getMostRecentFile(
            files_route, lambda x: [v for v in x if "xls" in v.lower()])

        total_finded = runMacro("modulo.ValidarACT", [temp_file_act, act[8]]) #(value, message)

        if total_finded is not None or total_finded[0] is not None:
            act[5] = int(float(total_finded[0]))    
        elif total_finded[0] is None and total_finded[1] is not None:
            insertInLog(f"Error encontrado {total_finded[1]}")
        else:
            insertInLog(f"Error encontrado no se pudo validar el act {act[1]} del archivo {os.path.basename(temp_file_act)}")

        os.remove(temp_file_act)
    else:
        _message = f"Error al buscar los registros del act {act[1]}, fecha: {_today.strftime('%m/%d/%Y')} - {_date_before.strftime('%m/%d/%Y')}"
        if elementIsVisible(driver, "//div[contains(@class,'mensaje')]", By.XPATH):
            validateErrorMessage(driver.find_element(
                By.XPATH, "//div[contains(@class,'mensaje')]"), _message)
        else:
            driver.get_screenshot_as_file(os.path.join(
                errors_route, f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"))
        pymsgbox.alert(_message)

    #volver a kardex
    chrome_driver.find_element(By.XPATH, f"//input[contains(@type,'submit') and contains(@value,'Volver')]").click()

    return act

if __name__ == "__main__":
    chrome_driver = webdriver.Chrome(
        executable_path=ChromeDriverManager().install(), options=chromeOptions)
    try:
        createNecesaryFolders(fr"{os.path.dirname(os.path.realpath(__file__))}", ["files", "_errors"])
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
        income_file = None
        if _result:
            income_file = getMostRecentFile(
                files_route, lambda x: [v for v in x if "xls" in v.lower()])

        # Buscar Detalles saldos
        returnHomeFrame(chrome_driver, True)
        chrome_driver.switch_to.frame("izquierda")
        search(chrome_driver, "detalles saldos")

        # buscar en el frame central
        returnHomeFrame(chrome_driver, True)
        chrome_driver.switch_to.frame("central")

        # descargar archivo saldos
        _result = downloadBalanceFile(chrome_driver)
        balance_file = None
        if _result:
            balance_file = getMostRecentFile(
                files_route, lambda x: [v for v in x if "xls" in v.lower()])

        chrome_driver.close()

        path_init = '\\'.join(initiatives_file.split('\\')[:-1])
        # necesario para las formulas de excel
        path_end = '\\' + "[" + initiatives_file.split('\\')[-1] + "]"
        initiatives_file = fr"{path_init}{path_end}"

        # ejecutar macro eliminar iniciativas completas
        runMacro('modulo.EliminarIniciativasCompletas')
        # ejecutar macro para actualizar iniciativas
        runMacro('modulo.ActualizarIniciativas', [initiatives_file])

        # actualizar los ingresos
        if income_file is not None:
            runMacro('modulo.ActualizarIngresos', [income_file])

        # actualizar los saldos (Inventario)
        if balance_file is not None:
            runMacro('modulo.ActualizarInventarios', [balance_file])

        # actualizar los balances (Ingresos y saldo)
        runMacro('modulo.ActualizarBalances')

        #Obtener acts
        acts = runMacro('modulo.ObtenerACTs')
        # Buscar kardex
        returnHomeFrame(chrome_driver, True)
        chrome_driver.switch_to.frame("izquierda")
        search(chrome_driver, "Kardex")

        #obtener asignaciones actualizadas
        temp_asigments = validateACTS(chrome_driver, acts)
        
        #insertar las asignaciones actualizadas
        runMacro("modulo.insertarACTS", [temp_asigments])

        pymsgbox.alert("\n Proceso Terminado, ya puede cerrar la ventana \n")
        print("\n Proceso Terminado, ya puede cerrar la ventana \n")

    except Exception as e:
        exception(e)
