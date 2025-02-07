# main.py
import sys
import re
import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException, NoSuchElementException


from flujo_correo import flujo_correo
from docx_convert import process_main
from rad_utils import get_rad_number, numero_radicado_global

def main():
    service = Service("C:/Users/SBURGOSP/EdgeDriver/edgedriver_win64/msedgedriver.exe")
    edge_options = Options()
    driver = webdriver.Edge(service=service, options=edge_options)

    try:
        # 1) Abrir la plataforma Mercurio
        driver.get("https://epm-vapp47.epm.com.co/mercurio/index.jsp")
        wait = WebDriverWait(driver, 20)
        
        # Guardar el identificador de la ventana principal
        ventana_main = driver.current_window_handle

        # Proceso de login
        input_usuario = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.NAME, 'usuario'))
        )
        input_login = driver.find_element(By.NAME, 'contrasena')
        boton_login = driver.find_element(By.NAME, 'Submit')

        input_usuario.send_keys('RPATBot01.cens')
        input_login.send_keys('Dz5npqa+')
        boton_login.click()

        # Esperar y ubicar el menú principal
        menu_principal = WebDriverWait(driver, 35).until(
            EC.presence_of_element_located((By.ID, "Bar4"))
        )

        # Moverse al menú principal
        accion_mover = ActionChains(driver)
        accion_mover.move_to_element(menu_principal).perform()

        # Hacer clic en la opción de búsqueda avanzada
        submenu_opcion = WebDriverWait(driver, 35).until(
            EC.visibility_of_element_located((By.ID, "menuItem4_6"))
        )
        submenu_opcion.click()

        # Rellenar parámetros de búsqueda
        numRadicado_input = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.NAME, 'idDocumento'))
        )
        checkBox_form = driver.find_element(By.NAME, 'activarFecha')
        boton_aceptar = driver.find_element(By.NAME, 'aceptar')
        
        numRadicado_input.send_keys('20251020002024')  
        checkBox_form.click()
        boton_aceptar.click()

        # Validar la columna "Origen"
        columna_radicado_origen = '//*[@id="tablaBusquedaAvanzada"]/tbody/tr/td[5]'
        columna_element = WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.XPATH, columna_radicado_origen))
        )
        columna_texto = columna_element.text.strip()

        numero_radicado = get_rad_number(driver)

        # Es buena práctica verificar si obtuvimos el número correctamente
        if numero_radicado:
            print(f"Número de radicado obtenido exitosamente: {numero_radicado}")
        else:
            print("No se pudo obtener el número de radicado")


        # Determinar el flujo según el contenido de la columna
        if columna_texto == "00000000-0000-0000-0000-000000000000":
            print("El flujo corresponde a página web.")

        elif columna_texto.isdigit() and 3 <= len(columna_texto) <= 4:
            print("El flujo corresponde a correo electrónico.")
            # Ejecutar el flujo de correo
            if flujo_correo(driver, ventana_main):
                print("Iniciando proceso de conversión de documentos...")
                # AQUÍ se llama la función main() de docx_convert
                process_main()
            else:
                print("Error durante la ejecución del flujo de correo.")
        
        elif columna_texto.isdigit() and len(columna_texto) == 14:
            print("El flujo corresponde a oficina.")
        elif not columna_texto:
            print("El flujo corresponde a oficina.")
        else:
            print(f"Contenido inesperado en la columna: {columna_texto}")

    except Exception as e:
        print(f"Se produjo un error en main: {str(e)}")
        
    finally:
        input("Presiona Enter para cerrar el navegador...")
        if 'driver' in locals():
            driver.quit()

if __name__ == "__main__":
    main()
