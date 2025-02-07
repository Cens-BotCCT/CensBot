# En rad_utils.py
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

numero_radicado_global = None  # Inicializamos la variable global

def get_rad_number(driver, timeout=10):
    global numero_radicado_global  # Declaramos que usaremos la variable global
    try:
        xpath = '//*[@id="tablaBusquedaAvanzada"]/tbody/tr/td[2]/a'
        elemento = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.XPATH, xpath))
        )
        
        # Actualizamos la variable global
        numero_radicado_global = elemento.text.strip()
        
        # Agregamos logs para verificar
        print(f"Número de radicado obtenido: {numero_radicado_global}")
        print(f"Tipo de dato: {type(numero_radicado_global)}")
        
        return numero_radicado_global
    except Exception as e:
        print(f"Error al obtener número de radicado: {str(e)}")
        return None