# flujo_correo.py
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Función para manejar el flujo de correo electrónico
def flujo_correo(driver, ventana_main):
    """
    Maneja el flujo de correo electrónico en la plataforma Mercurio.
 
    Parámetros:
      - driver: Instancia del WebDriver de Selenium.
      - ventana_main: Identificador de la ventana principal.
    """
    try:
        # Verifica que solo haya una ventana abierta inicialmente
        assert len(driver.window_handles) == 1
 
        for i in range(4, 6):  # Itera sobre los selectores (puedes ajustar el rango según necesites)
            selector = f'#tablaBusquedaAvanzada > tbody > tr > td.sorting_1 > a:nth-child({i}) > img'
 
            try:
                driver.find_element(By.CSS_SELECTOR, selector).click()
 
                wait = WebDriverWait(driver, 10)
                wait.until(EC.number_of_windows_to_be(2))
 
                # Cambiar a la nueva ventana
                ventana_dos = None
                for handle in driver.window_handles:
                    if handle != ventana_main:
                        ventana_dos = handle
                        driver.switch_to.window(ventana_dos)
                        break
 
                try:
                    wait.until(EC.presence_of_element_located((By.ID, 'barramensaje')))
 
                    ver_anexo = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, '//*[@id="scroll"]/table/tbody/tr[2]/td[1]/a'))
                    )
                    ver_anexo.click()
 
                    wait.until(EC.number_of_windows_to_be(3))
 
                    time.sleep(5)
 
                    driver.switch_to.window(ventana_dos)
                    driver.close()
                    driver.switch_to.window(ventana_main)
 
                    break  # Si se encontró el anexo, salir del bucle
 
                except Exception:
                    driver.switch_to.window(ventana_dos)
                    driver.close()
                    driver.switch_to.window(ventana_main)
                    continue
 
            except Exception as e:
                print(f"Error al interactuar con el enlace {i}: {e}")
                continue
 
        boton_salir = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, 'Bar7'))
        )
        boton_salir.click()
 
        return True
 
    except Exception as e:
        print(f"Error en flujo_correo: {e}")
        return False