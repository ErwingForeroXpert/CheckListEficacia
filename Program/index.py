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
        chrome_driver.close()

        most_recent_file = getMostRecentFile(
            files_route, lambda x: [v for v in x if "xls" in v.lower()])
        path_init = '\\'.join(most_recent_file.split('\\')[:-1])
        # necesario para las formulas de excel
        path_end = '\\' + "[" + most_recent_file.split('\\')[-1] + "]"
        most_recent_file = fr"{path_init}{path_end}"

        book_checklist = None
        parent_folder = os.path.join(os.getcwd(), "Files")
        for fname in os.listdir(parent_folder):
            if "xlsm" in fname and "~" not in fname:
                book_checklist = xw.Book(os.path.join(parent_folder, fname))
                break

        if book_checklist is None:
            raise Exception("Book with macro not found or open")

        # ejecutar macro eliminar iniciativas completas
        book_checklist.macro('Módulo1.EliminarIniciativasCompletas')()

        # ejecutar macro para actualizar iniciativas
        book_checklist.macro('Módulo1.ActualizarIniciativas')(most_recent_file)
        book_checklist.save()
        book_checklist.close()

        print("\n Proceso Terminado, ya puede cerrar la ventana \n")

    except Exception as e:
        exception(e)