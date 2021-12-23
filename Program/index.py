from utils import *
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
documents_route = fr"{os.path.dirname(os.path.realpath(__file__))}\documents"
errors_route = fr"{os.path.dirname(os.path.realpath(__file__))}\_errors"

# insertar la direccion de descarga
chromeOptions = webdriver.ChromeOptions()
prefs = {"download.default_directory": files_route}
chromeOptions.add_experimental_option("prefs", prefs)
chromeOptions.add_argument("start-maximized")


#@exceptionHandler
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


#@exceptionHandler
def login(driver):
    user_element = driver.find_element_by_xpath("//input[@id='username']")
    user_element.send_keys(config["USER"])

    pass_element = driver.find_element_by_xpath("//input[@id='password']")
    pass_element.send_keys(config["PASSWORD"])
    pass_element.send_keys(Keys.ENTER)


#@exceptionHandler
def runMacro(nameMacro, _args=None):
    # ejecutar macro
    result = None
    parent_folder = documents_route
    _book = None
    for fname in os.listdir(parent_folder):
        if "xlsm" in fname and "~" not in fname:
            waitBookDisable(os.path.join(parent_folder, fname))
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


#@exceptionHandler
def returnHomeFrame(driver, switch_default=False):
    if switch_default:
        driver.switch_to.default_content()
    waitElement(
        driver,
        "//frame[@name='mainFrame']",
        By.XPATH)
    driver.switch_to.frame("mainFrame")


#@exceptionHandler
def click_option(_select, option, precision="same"):
    options_type_document = Select(_select).options
    for _op in options_type_document:
        validation = _op.text.lower() == option.lower(
        ) if precision == "same" else option.lower() in _op.text.lower()
        if validation:
            _op.click()
            return True
    return False


#@exceptionHandler
def validateErrorMessage(element, extra_info=""):
    if foundInErrorMessages(element.text):
        insertInLog(f"Error encontrado {element.text} {extra_info}")
        return True
    else:
        return False


#@exceptionHandler
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
        driver.find_element(
            By.XPATH, "//a[contains(@href,'tipo_exportacion=XLS')]").click()
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


#@exceptionHandler
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
        driver.find_element(
            By.XPATH, "//a[contains(@href,'tipo_exportacion=XLS')]").click()
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


#@exceptionHandler
def validateACTS(driver, registers):
    temp_registers = []

    returnHomeFrame(driver, True)
    driver.switch_to.frame("central")
    waitElement(driver, "//table[@class='tablaexhibir']", By.XPATH)
    driver.find_element_by_xpath(
        "//table[@class='tablaexhibir TablaContainerTable']/tbody/tr/td[@class='td2']//a[contains(text(), 'Kardex')]").click()

    for register in registers:
        if "act" in str(register[2]).lower() and ifErrorFalse(lambda x: int(x) < 100, register[19]):
            if "bgta" in str(register[1]).lower():
                temp_act = validateValueACT(driver, register, by="excel")
            else:
                temp_act = validateValueACT(driver, register)
        else:
            temp_act = list(register)

        # first 9 registers [0,1,..] index
        temp_registers.append(tuple(temp_act[:9]))

    return tuple(temp_registers)


#@exceptionHandler
def validateValueACT(driver, act, by="table"):
    temp_act = list(act)
    # seelct option selector
    returnHomeFrame(driver, True)
    driver.switch_to.frame("central")
    if not click_option(driver.find_element(By.XPATH,
                                            "//select[contains(@name,'punto')]"), temp_act[1].split(" ")[0], "similar"):
        insertInLog(
            f"Error al buscar la opcion {temp_act[1].split(' ')[0]} act {temp_act[2]}", "error")
        return temp_act

    _today = datetime.now()
    # margin of days for search act information
    _date_before = _today - timedelta(days=25)
    driver.find_element_by_name("articulo").clear()
    driver.find_element_by_name("articulo").send_keys(
        temp_act[2])  # insert act
    driver.find_element_by_name("fecha_final").clear()
    driver.find_element_by_name(
        "fecha_final").send_keys(_today.strftime('%m/%d/%Y'))  # insert final date
    driver.find_element_by_name("fecha_inicial").clear()
    driver.find_element_by_name("fecha_inicial").send_keys(  # insert init date
        _date_before.strftime('%m/%d/%Y'))
    driver.find_element(
        By.XPATH, "//input[contains(@type,'submit') and contains(@value,'Continuar')]").click()

    # wait for the button "Continuar" to disappear
    returnHomeFrame(driver, True)
    driver.switch_to.frame("central")
    waitElementDisable(
        driver, "//input[contains(@type,'submit') and contains(@value,'Continuar')]", By.XPATH)
    time.sleep(2)

    # validate if the table of records exist and is searched specifically by table
    if elementIsVisible(driver, "//table[contains(@class,'tablaexhibir')]", By.XPATH, wait=2) and by == "table":
        len_table = len(driver.find_elements(
            By.XPATH, "//table[contains(@class,'tablaexhibir')]/tbody/tr"))
        if len_table > 2:
            total_finded = driver.find_element(
                By.XPATH, f"//table[contains(@class,'tablaexhibir')]/tbody/tr[{len_table}]/th[16]").text

            temp_act[5] = max(string2Number(temp_act[5]), string2Number(total_finded)) if ifErrorFalse(
                string2Number, total_finded) else string2Number(total_finded)

    elif elementIsVisible(driver, "//a[contains(@href,'tipo_exportacion=XLS')]", By.XPATH) or elementIsVisible(driver, "//a/small[contains(text(),'Descargar Archivo')]", By.XPATH):
        # if first selector isn't found the following is used
        if elementIsVisible(driver, "//a[contains(@href,'tipo_exportacion=XLS')]", By.XPATH):
            driver.find_element(
                By.XPATH, "//a[contains(@href,'tipo_exportacion=XLS')]").click()
        else:
            driver.find_element(
                By.XPATH, "//a/small[contains(text(),'Descargar Archivo XLS')]").click()

        waitDownload(files_route)
        temp_file_act = getMostRecentFile(
            files_route, lambda x: [v for v in x if "xls" in v.lower()])

        # run macro for validate the info in file downloaded
        total_finded = runMacro("modulo.ValidarACT", [
                                temp_file_act, temp_act[8]])  # (value, message)

        if isinstance(total_finded, (list, tuple)) and (not isEmpty(total_finded[0]) and isEmpty(total_finded[1])):
            temp_act[5] = max(string2Number(temp_act[5]), string2Number(total_finded[0])) if ifErrorFalse(
                string2Number, total_finded[0]) else string2Number(total_finded[0])
        elif isinstance(total_finded, (list, tuple)) and not isEmpty(total_finded[1]):
            insertInLog(f"{temp_act[1]}-{temp_act[2]}", "error")
        else:
            insertInLog(
                f"no se pudo validar el act {temp_act[1]}-{temp_act[2]} del archivo {os.path.basename(temp_file_act)}", "error")

        os.remove(temp_file_act)
    else:
        # if all of the above doesn't work
        _message = f"al buscar los registros del act: {temp_act[1]} - {temp_act[2]}, fecha {_date_before.strftime('%m/%d/%Y')} - {_today.strftime('%m/%d/%Y')}"
        if elementIsVisible(driver, "//div[contains(@class,'mensaje')]", By.XPATH):
            validateErrorMessage(driver.find_element(
                By.XPATH, "//div[contains(@class,'mensaje')]"), _message)
        else:
            driver.get_screenshot_as_file(os.path.join(
                errors_route, f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"))
        # pymsgbox.alert(_message)

    # return to kardex
    if elementIsVisible(driver, "//input[contains(@type,'submit') and contains(@value,'Volver')]", By.XPATH):
        driver.find_element(
            By.XPATH, "//input[contains(@type,'submit') and contains(@value,'Volver')]").click()
    else:
        # search kardex
        returnHomeFrame(driver, True)
        driver.switch_to.frame("izquierda")
        search(driver, "Kardex")
        returnHomeFrame(driver, True)
        driver.switch_to.frame("central")
        waitElement(driver, "//table[@class='tablaexhibir']", By.XPATH)
        driver.find_element_by_xpath(
            "//table[@class='tablaexhibir TablaContainerTable']/tbody/tr/td[@class='td2']//a[contains(text(), 'Kardex')]").click()

    return temp_act


if __name__ == "__main__":
    chrome_driver = webdriver.Chrome(
        executable_path=ChromeDriverManager().install(), options=chromeOptions)
    try:
        createNecesaryFolders(fr"{os.path.dirname(os.path.realpath(__file__))}", [
                              "files", "_errors"])
        deleteTemporals(files_route)
        chrome_driver.get(config["URL_EFICACIA"])

        # wait login
        returnHomeFrame(chrome_driver)

        waitElement(chrome_driver, "//input[@id='username']", By.XPATH, True)
        login(chrome_driver)

        if elementIsVisible(chrome_driver, "//div[@class='error login_error' and contains(text(), 'YA TIENE SESION ACTIVA')]", By.XPATH):
            close_sessions_element = chrome_driver.find_element_by_xpath(
                "//input[@type='checkbox'][@name='cerrar_sesiones_anteriores']")
            chrome_driver.execute_script(
                "arguments[0].setAttribute('checked',arguments[1])", close_sessions_element, True)
            login(chrome_driver)

        # wait home of the page
        waitElement(
            chrome_driver,
            "//div[@id='Layer1']",
            By.XPATH)

        # Look for the technical sheet
        waitElement(
            chrome_driver,
            "//frame[@name='izquierda']",
            By.XPATH)
        chrome_driver.switch_to.frame("izquierda")
        search(chrome_driver, "ficha")

        returnHomeFrame(chrome_driver, True)

        # look for the frame central
        waitElement(
            chrome_driver,
            "//frame[@name='central']",
            By.XPATH)
        chrome_driver.switch_to.frame("central")

        waitElement(chrome_driver, "//table[@class='tablaexhibir']", By.XPATH)
        chrome_driver.find_element_by_xpath(
            "//table[@class='tablaexhibir TablaContainerTable']/tbody/tr/td[@class='td2'][@align='left']/font").click()

        # search the articles
        waitElement(chrome_driver,
                    "//table/b[contains='*Son Campos Obligatorios']")
        chrome_driver.find_element_by_xpath(
            "//table//input[@type='submit']").click()

        # download the file
        waitElementDisable(
            chrome_driver, "//table//input[@type='submit']", By.XPATH)
        chrome_driver.find_element_by_xpath(
            "//a/small[contains(text(),'Archivo XLS')]").click()
        time.sleep(1)
        waitDownload(files_route)
        time.sleep(1)

        initiatives_file = getMostRecentFile(
            files_route, lambda x: [v for v in x if "xls" in v.lower()])

        # look for document details
        returnHomeFrame(chrome_driver, True)
        chrome_driver.switch_to.frame("izquierda")
        search(chrome_driver, "Detalles Documentos")

        # look for the central frame
        returnHomeFrame(chrome_driver, True)
        chrome_driver.switch_to.frame("central")

        # download income file
        _result = downloadIncomeFile(chrome_driver)
        time.sleep(1)
        income_file = None
        if _result:
            income_file = getMostRecentFile(
                files_route, lambda x: [v for v in x if "xls" in v.lower()])

        # look for the balance file
        returnHomeFrame(chrome_driver, True)
        chrome_driver.switch_to.frame("izquierda")
        search(chrome_driver, "detalles saldos")

        # look for the central frame
        returnHomeFrame(chrome_driver, True)
        chrome_driver.switch_to.frame("central")

        # download balance file
        _result = downloadBalanceFile(chrome_driver)
        time.sleep(1)
        balance_file = None
        if _result:
            balance_file = getMostRecentFile(
                files_route, lambda x: [v for v in x if "xls" in v.lower()])


        path_init = '\\'.join(initiatives_file.split('\\')[:-1])
        # necesario para las formulas de excel
        path_end = '\\' + "[" + initiatives_file.split('\\')[-1] + "]"
        initiatives_file = fr"{path_init}{path_end}"

        # Execute delete complete initiatives
        runMacro('modulo.EliminarIniciativasCompletas')
        # ejecutar macro para actualizar iniciativas
        runMacro('modulo.ActualizarIniciativas', [initiatives_file])

        # update incomes
        if income_file is not None:
            runMacro('modulo.ActualizarIngresos', [income_file])

        # actualizar los saldos (Inventario)
        if balance_file is not None:
            runMacro('modulo.ActualizarInventarios', [balance_file])

        # # actualizar los balances (Ingresos y saldo)
        runMacro('modulo.ActualizarBalances')

        # Obtener acts
        acts = runMacro('modulo.ObtenerACTs')
        # Buscar kardex
        returnHomeFrame(chrome_driver, True)
        chrome_driver.switch_to.frame("izquierda")
        search(chrome_driver, "Kardex")

        # obtener asignaciones actualizadas
        temp_asigments = validateACTS(chrome_driver, acts)
        chrome_driver.close()

        # insertar las asignaciones actualizadas
        runMacro("modulo.insertarACTS", [temp_asigments])

        pymsgbox.alert("\n Proceso Terminado, ya puede cerrar la ventana \n")
        print("\n Proceso Terminado, ya puede cerrar la ventana \n")

    except Exception as e:
        exception(e)
