from robocorp.tasks   import task, teardown
from robocorp         import browser
from RPA.HTTP         import HTTP
from RPA.PDF          import PDF
# from RPA.Excel.Files  import Files

import os
import pandas as pd
import zipfile

import time

# Definimos aquí las variables globales. Podrían ser también parámetros en un fichero de configuración
url_order_website  = "https://robotsparebinindustries.com/#/robot-order"
url_order_csv_file = "https://robotsparebinindustries.com/orders.csv"
download_directory = "descargas"

def open_robot_order_website():
    """ Abriremos la página de pedidos"""
    browser.configure(browser_engine="chrome", slowmo=100)
    browser.goto(url_order_website)

    # Esperamos a que aparezca el mensaje de la ventana emergente y pulsamos el botón OK
    pagina = browser.page()
    pagina.wait_for_selector('//button[.="OK"]', state='visible', timeout=30000, strict=False)
    pagina.click('//button[.="OK"]')

def get_orders():
    """ Desde aquí se llamará a la URL que descargará el CSV de los pedidos y
        se guardará en una variable que se devolverá al llamador.
    """
    # Definimos un directorio de descarga.
    # En este caso, el directorio de descarga será el directorio de trabajo más el subdirectorio descargas.
    execution_directory = os.getcwd()
    complete_download_directory = os.path.join(execution_directory, download_directory)

    # Descargamos el fichero
    http_descarga = HTTP()
    http_descarga.download(url_order_csv_file, complete_download_directory, overwrite=True)

    # Cargamos el fichero en una variable y la devolvemos
    orders_csv = pd.read_csv(complete_download_directory + "/orders.csv")
    return orders_csv

def fill_the_form(orders):
    """ Dado el texto del CSV de pedidos, rellenaremos el formulario de pedido"""
    # Creamos un objeto de la página
    pagina = browser.page()
    
    for indice, orden in orders.iterrows():
        # Rellenamos el formulario
        pagina.locator("select#head").select_option(value=str(orden["Head"]))
        selector = "input#id-body-" + str(orden["Body"])
        pagina.locator(selector).set_checked(True)
        pagina.locator("//input[contains(@placeholder, 'Enter the part')]").fill(str(orden["Body"]))
        pagina.locator("input#address").fill(str(orden["Address"]))
        pagina.locator("button#order").click()

        def order_charged():
            # Esperamos a que se cargue la orden y capturamos el resultado.
            pagina.wait_for_selector('#order-another', state='visible', timeout=1000, strict=False)
            capture_order()

        def capture_order():
            # Realizamos una captura de pantalla con el resultado de la orden.
            pagina_resultado = browser.page()
            nombre_fichero_jpg = os.getcwd() + "/" + download_directory + "/order_" + str(indice + 1) + ".jpg"
            pagina_resultado.screenshot(path=nombre_fichero_jpg, type="jpeg")
            
            # También guardamos el resultado en un fichero PDF.
            order_HTML = pagina_resultado.locator("div#order-completion").inner_html()
            nombre_fichero_pdf = os.getcwd() + "/" + download_directory + "/order_" + str(indice + 1) + ".pdf"
            fichero_PDF = PDF()
            fichero_PDF.html_to_pdf(order_HTML, nombre_fichero_pdf)

            # Incluiremos la imagen dentro del PDF.
            lista_ficheros = [nombre_fichero_pdf, nombre_fichero_jpg]
            fichero_PDF.add_files_to_pdf(files=lista_ficheros, target_document=nombre_fichero_pdf)

            # Eliminamos la imagen que ya está incluida en el PDF.
            os.remove(nombre_fichero_jpg)

        def wait_until_order_succeeds(intentos, intervalo):
            for indice in range(intentos):
                try:
                    order_charged()
                    break
                except:
                    time.sleep(intervalo)
                    pagina.locator("button#order").click()
            else:
                raise Exception(f"Error al introducir la orden")

        wait_until_order_succeeds(5, 1)
           
        # Pulsamos para cargar una nueva orden, esperamos a que salga la ventana emergente y damos al OK
        pagina.locator('#order-another').click()
        pagina.wait_for_selector('//button[.="OK"]', state='visible', timeout=30000, strict=False)
        pagina.click('//button[.="OK"]')

def archive_receipts():
    """ Metemos todos los ficheros PDF en un fichero ZIP """
    # Si el fichero zip ya existe, lo eliminamos primero.
    if os.path.exists(os.getcwd() + "/" + download_directory + "/receipts.zip"):
        os.remove(os.getcwd() + "/" + download_directory + "/receipts.zip")

    # Recuperamos todos los archivos generados en el directorio de descarga 
    # y aquellos que sean PDF los metemos en un fichero ZIP.
    # El resto de ficheros los eliminamos.
    execution_directory = os.getcwd() + "/" + download_directory
    for carpeta, subcarpeta, nombrefichero in os.walk(execution_directory):
        for fichero in nombrefichero:
            nombre_fichero = os.path.join(carpeta, fichero)
            if fichero.endswith(".pdf"):
                zipfile.ZipFile(download_directory + "/receipts.zip", "a").write(nombre_fichero, fichero)
            os.remove(nombre_fichero)

@task
def order_robots_from_RobotSpareBin():
    open_robot_order_website()
    orders = get_orders()
    fill_the_form(orders) 
    archive_receipts()

@teardown
def end_tasks(tarea):
    print(f"Finalizando la tarea {tarea}")
    