from utils import waitDownload, waitElement, waitElementDisable, deleteTemporals, elementIsVisible, getMostRecentFile
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from dotenv import dotenv_values
from selenium import webdriver
import xlwings as xw
from logging import exception
import time
import os


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
        raise Exception(f"en la ruta {parent_folder}, el libro no se encuentra o no se puede abrir")
    
    result = _book.macro(nameMacro)(*_args) if _args is not None else _book.macro(nameMacro)()
    
    _book.save()

    if len(_book.app.books) == 1:
        _book.app.quit()
    else:
        _book.close()

    return result

if __name__ == "__main__":
    chrome_driver = webdriver.Chrome(
        executable_path=ChromeDriverManager().install(), options=chromeOptions)
    try:
        deleteTemporals(files_route)
        chrome_driver.get(config["URL_EFICACIA"])

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
        chrome_driver.switch_to.frame("izquierda")
        search(chrome_driver, "ficha")

        # buscar en el frame central
        chrome_driver.switch_to.default_content()
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

        Initiatives_file = getMostRecentFile(
            files_route, lambda x: [v for v in x if "xls" in v.lower()])
        
        # Buscar Detalles documentos
        chrome_driver.switch_to.frame("izquierda")
        search(chrome_driver, "Detalles Documentos")

        # buscar en el frame central
        chrome_driver.switch_to.default_content()
        chrome_driver.switch_to.frame("central")

        waitElement(chrome_driver, "//table[@class='tablaexhibir']", By.XPATH)
        chrome_driver.find_element_by_xpath(
            "//table[@class='tablaexhibir TablaContainerTable']/tbody/tr/td[@class='td2'][@align='left']/font").click()
            
        chrome_driver.close()

        path_init = '\\'.join(Initiatives_file.split('\\')[:-1])
        # necesario para las formulas de excel
        path_end = '\\' + "[" + Initiatives_file.split('\\')[-1] + "]"
        Initiatives_file = fr"{path_init}{path_end}"

        # ejecutar macro eliminar iniciativas completas
        runMacro('Módulo1.EliminarIniciativasCompletas')
        # ejecutar macro para actualizar iniciativas
        runMacro('Módulo1.ActualizarIniciativas', [Initiatives_file])


        print("\n Proceso Terminado, ya puede cerrar la ventana \n")

    except Exception as e:
        exception(e)