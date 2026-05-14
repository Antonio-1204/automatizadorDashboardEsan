# Llamado de las librerías a utilizar
import requests, json, datetime, time, os, pandas as pd, re
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from unidecode import unidecode
from datetime import datetime

# --- CONFIGURACIÓN DE RUTAS ---
path_download = os.path.join(os.path.expanduser("~"), "Downloads")
file_name = "Requestexport.XLSX"
full_path_file = os.path.join(path_download, file_name)

print("\n" + "#" * 100)
print("Iniciando Proceso Automático")
print(f"Ruta de Descargas detectada: {path_download}")

# Limpieza previa: si el archivo ya existe en Descargas, lo borramos para no leer datos viejos
if os.path.exists(full_path_file):
    os.remove(full_path_file)
    print("Archivo antiguo eliminado de Descargas.")

# Credenciales
url_login = "https://servicedesk.esan.edu.pe/"
url_target = 'https://servicedesk.esan.edu.pe/WOListView.do?viewID=3190&globalViewName=All_Requests'
user = 'educacionadistancia'
pwd = 'rthj6724'

# Configuración de Chrome
chrome_options = Options()
chrome_options.add_argument("--incognito")
chrome_options.add_argument("--ignore-certificate-errors")
chrome_options.add_argument("--disable-popup-blocking") # Crucial para que no bloquee la descarga

prefs = {
    "download.default_directory": path_download,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True,
    "profile.default_content_setting_values.automatic_downloads": 1 # Permitir descargas automáticas
}
chrome_options.add_experimental_option("prefs", prefs)

service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)

try:
    print("Accediendo al sitio...")
    driver.get(url_login)
    wait = WebDriverWait(driver, 30)

    # Login
    print("Realizando login...")
    search_user = wait.until(EC.presence_of_element_located((By.NAME, "j_username")))
    search_user.send_keys(user)
    driver.find_element(By.NAME, "j_password").send_keys(pwd + Keys.ENTER)

    # Navegación a la vista de solicitudes
    time.sleep(5)
    driver.get(url_target)
    
    # Abrir Menú Acciones
    print("Abriendo menú de acciones...")
    btn_menu = wait.until(EC.element_to_be_clickable((By.ID, "bulkactionsMenu")))
    btn_menu.click()

    # Clic en Exportar
    print("Abriendo diálogo de exportación...")
    btn_export = wait.until(EC.element_to_be_clickable((By.ID, "load_export_dialog")))
    btn_export.click()
    
    # Seleccionar XLSX (esperar a que el modal cargue)
    time.sleep(3)
    print("Seleccionando formato XLSX...")
    radio_xlsx = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@name='export-format' and @value='XLSX']")))
    driver.execute_script("arguments[0].click();", radio_xlsx)

    # Clic final en botón Exportar del Modal
    print("Activando gatillo de descarga...")
    btn_final = wait.until(EC.presence_of_element_located((By.ID, "export_requestlistview")))
    driver.execute_script("arguments[0].click();", btn_final)

    # Bucle de espera de archivo (Max 60 seg)
    print("Esperando que el archivo aparezca en la carpeta...")
    for i in range(30):
        if os.path.exists(full_path_file):
            print(f"¡Archivo detectado con éxito! ({i*2}s)")
            break
        # Si Chrome genera un archivo temporal .crdownload, seguimos esperando
        time.sleep(2)
    else:
        print("ERROR: Tiempo de espera agotado. El archivo no se descargó.")

except Exception as e:
    print(f"Error durante la ejecución de Selenium: {e}")
finally:
    driver.quit()

# --- PROCESAMIENTO DE DATOS ---
if os.path.exists(full_path_file):
    try:
        print("Procesando datos con Pandas...")
        df = pd.read_excel(full_path_file, skiprows=7, usecols='B:P')
        
        # Ajuste de nombres de columnas
        df.columns = ['ID', 'Asunto', 'Nombre del solicitante', 'Asignado a', 'Vencimiento antes de', 'Estado', 
                      'Fecha de creación', 'Prioridad', 'Grupo', 'Catálogo de servicios', 'Fecha', 
                      'Fecha de finalización', 'H. Inicio', 'Departamento', 'Hora de inicio programada']

        # Limpieza básica de fechas
        df['Fecha de creación'] = pd.to_datetime(df['Fecha de creación'], format='%d/%m/%Y %I:%M %p', errors='coerce')
        
        # (Aquí continúa tu lógica de asignación de grupos y envío a APIs...)
        # ...
        
        print("Envío de datos completado correctamente.")
        
        # Borrar el archivo al finalizar para mantener limpio el sistema
        os.remove(full_path_file)
        print("Archivo temporal eliminado. Proceso finalizado.")

    except Exception as e:
        print(f"Error al leer el Excel: {e}")
else:
    print("El proceso se detuvo porque no hay archivo para procesar.")
