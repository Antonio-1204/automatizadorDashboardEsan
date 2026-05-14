# Llamado de las librerías a utilizar
import requests, json, datetime, time, os, pyautogui, io, pandas as pd, re
import pygetwindow as gw
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
# --- NUEVOS IMPORTS PARA CORREGIR EL ERROR ---
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
# ---------------------------------------------
from unidecode import unidecode
from datetime import datetime

# --- CONFIGURACIÓN DE RUTAS AUTOMÁTICA ---
# Detecta la carpeta de descargas del usuario actual de Windows
path_download = os.path.join(os.path.expanduser("~"), "Downloads")
file_name = "Requestexport.XLSX"
full_path_file = os.path.join(path_download, file_name)

print("\n\n\n")
print("#" * 100)
print("Iniciando el proceso de llenado de datos")
print(f"Carpeta de trabajo (Descargas): {path_download} \n\n")

file_dwnload = True

# Verificación en la carpeta de Descargas
if file_name in os.listdir(path_download):
    print(f"Se encontró el archivo de solicitudes en Descargas - '{file_name}'")
    resp = input("Desea continuar con el archivo encontrado?  ([Y]yes/[N]no): ")
    print("#" * 100, end="\n\n")

    if (resp.lower() == 'y' or resp.lower() == 'yes'):
        print("Continuando con el archivo existente")
        file_dwnload = False
    else:
        file_dwnload = True
        # Eliminar el archivo existente en Descargas
        os.remove(full_path_file)
        print("Se ha eliminado el archivo existente para descargar uno nuevo")

if file_dwnload:
    print("Iniciando proceso de descarga automática...")
    print("#" * 100, end="\n\n")

    # Credenciales
    url = "https://servicedesk.esan.edu.pe/"
    user = 'educacionadistancia'
    pwd = 'rthj6724'

    # Configura las opciones para el navegador
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--incognito")  # Activa el modo incógnito
    
    # --- PREFERENCIAS DE DESCARGA AUTOMÁTICA ---
    prefs = {
        "download.default_directory": path_download, # Carpeta Descargas
        "download.prompt_for_download": False,       # No preguntar dónde guardar
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    chrome_options.add_experimental_option("prefs", prefs)
    
    # Inicializa el controlador de Chrome
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)

    try:
        driver.get(url)
        # Redirección forzada
        url_target = 'https://servicedesk.esan.edu.pe/WOListView.do?viewID=3190&globalViewName=All_Requests'
        time.sleep(4)

        # Zona de Login
        wait = WebDriverWait(driver, 25)
        search_user = wait.until(EC.presence_of_element_located((By.NAME, "j_username")))
        search_user.send_keys(user)
        search_pwd = driver.find_element(By.NAME, "j_password")
        search_pwd.send_keys(pwd)
        search_pwd.send_keys(Keys.ENTER)

        time.sleep(3)
        driver.get(url_target)
        time.sleep(5)

        # 1. Clic en el menú de acciones
        print("Abriendo menú de acciones...")
        menu_actions = wait.until(EC.element_to_be_clickable((By.ID, "bulkactionsMenu")))
        menu_actions.click()
                
        # 2. Clic en la opción exportar
        print("Esperando opción de exportar...")
        export_option = wait.until(EC.element_to_be_clickable((By.ID, "load_export_dialog")))
        export_option.click()
        
        time.sleep(2) 
        
        # 3. Seleccionar el formato XLSX específicamente
        print("Seleccionando formato XLSX...")
        radio_xlsx = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@name='export-format' and @value='XLSX']")))
        driver.execute_script("arguments[0].click();", radio_xlsx)

        time.sleep(0.5)

        # 4. Hacer clic en el botón de Exportar
        print("Activando descarga...")
        btn_exportar = wait.until(EC.presence_of_element_located((By.ID, "export_requestlistview")))
        driver.execute_script("arguments[0].click();", btn_exportar)

        # Espera hasta que el archivo aparezca en la carpeta de Descargas
        intentos = 0
        while (file_name not in os.listdir(path_download)):
            print(f"Esperando que el archivo aparezca en Descargas... ({intentos*2}s)")
            time.sleep(2)
            intentos += 1
            if intentos > 30: # Timeout de 60 segundos
                break

        print("Descarga del archivo realizada correctamente")

    except Exception as e:
        print(f'Ocurrió un error: {e}')
    finally:
        driver.quit()


# Zona de tratamiento de los datos
if file_name in os.listdir(path_download):
    # Lectura desde la ruta de Descargas
    df = pd.read_excel(full_path_file, skiprows = 7, usecols = 'B:P')

    # Ajuste del nombre de columnas
    df.columns = ['ID', 'Asunto',  'Nombre del solicitante','Asignado a', 'Vencimiento antes de', 'Estado' ,'Fecha de creación', 'Prioridad', 'Grupo', 'Catálogo de servicios', 'Fecha', 'Fecha de finalización', 'H. Inicio', 'Departamento','Hora de inicio programada']

    # Reordenando las columnas
    df = df[['ID', 'Asunto', 'Nombre del solicitante', 'Asignado a', 'Vencimiento antes de', 'Estado', 'Fecha de creación', 'Prioridad', 'Grupo', 'Catálogo de servicios', 'Fecha', 'Fecha de finalización', 'H. Inicio', 'Departamento','Hora de inicio programada']]

    # Reordenando las filas
    df = df.sort_values(by='ID')

    # Ajuste del tiempo y asignación
    df['Fecha de creación'] = pd.to_datetime(df['Fecha de creación'], format='%d/%m/%Y %I:%M %p')
    df.loc[df['Fecha de finalización'] == '-', 'Fecha de finalización'] = datetime.now().strftime('%d/%m/%Y %I:%M %p')
    df['Fecha de finalización'] = pd.to_datetime(df['Fecha de finalización'], format='%d/%m/%Y %I:%M %p')
    df.loc[df['Vencimiento antes de'] == '-', 'Vencimiento antes de'] = ''
    df['Vencimiento antes de'] = pd.to_datetime(df['Vencimiento antes de'], format='%d/%m/%Y %I:%M %p', errors='coerce')
    
    df_tf = df.copy()

    # Asignación de áreas
    df.loc[df['Asignado a'] == 'Martha Medina Cloke', 'Grupo'] = 'ESANLabs-Diseño'
    df.loc[df['Asignado a'] == 'Gabriel Juarez Postigo', 'Grupo'] = 'ESANLabs - Ted'
    df.loc[df['Asignado a'] == 'Maria Jose Sanchez Stein', 'Grupo'] = 'ESANLabs - Ted'
    df.loc[df['Asignado a'].str.contains('Andrea', case = False, na=False), 'Grupo'] = 'ESANLabs - Ted'
    df.loc[df['Asunto'].str.contains('Reenviar este correo', case = False, na=False), 'Grupo'] = 'Enlaces Zoom'
    df.loc[df['Asunto'].str.contains('CAP -', case = False, na=False), 'Grupo'] = 'CAP'
    df.loc[df['Asunto'].str.contains('AESPI - ', case = False, na=False), 'Grupo'] = 'AESPI'
    df.loc[(df['Asunto'].str.contains('zoom', case = False, na=False)) & (df['Grupo'] == 'ESANLabs-Multimedia'), 'Grupo'] = 'Enlaces Webinar' 
    df.loc[df['Asunto'].str.contains('PRY', case = False, na=False), 'Grupo'] = "Esanlabs - Proyectos"
    df.loc[df['Asignado a'] == 'Yessica Renee Guardamino Vera', 'Grupo'] = 'ESANLabs - Ted'
    
    df.loc[(df['Fecha de creación'] > '2024-08-01') & (df['Asignado a'].str.contains('Juan Carlos Leyva Carrasco', case = False, na=False)), 'Grupo'] = 'ESANLabs-Multimedia'
    df.loc[(df['Fecha de creación'] < '2024-08-01') & (df['Asignado a'].str.contains('Juan Carlos Leyva Carrasco', case = False, na=False)), 'Grupo'] = 'ESANLabs - Ted'

    # Zona de Funciones
    def Api(action = 'Get', sheet = 'General' , data = [], filaInicio = 0, filaFin = 0 , url = ''):
        payload = {
            'action': action,
            'sheet': sheet,
            'data': data,
            'filaInicio': filaInicio,
            'filaFin':filaFin
        }
        try:
            response = requests.post(url, json=payload)
            return response.json()
        except Exception as e:
            print("Request failed:", e)

    def df_to_array(aux):
        aux = aux.copy()
        aux['Fecha de creación'] = aux['Fecha de creación'].astype(str)
        aux['Fecha de finalización'] = aux['Fecha de finalización'].astype(str)
        aux['Vencimiento antes de'] = aux['Vencimiento antes de'].astype(str)
        column_names = aux.columns.values.tolist()
        data_lists = aux.values.tolist()
        result = [column_names] + data_lists
        for x in range (len(result)):
            if result[x][4] == 'NaT':
                result[x][4] = '-'
        return result

    # Carga de información
    url_gs = 'https://script.google.com/macros/s/AKfycby-d58Qp5lvoLHvBALoIVBIDudbEnOgDrGY-V9CNcs2xSo6lufl6TmtcwnhrVBjnM4-Zw/exec'
    Api('Post','GENERAL',df_to_array(df), url=url_gs)
    Api('Post','MULTIMEDIA',df_to_array(df.loc[((df['Grupo'] == 'ESANLabs-Multimedia') | (df['Grupo'] == 'Enlaces Webinar')) & (df['Fecha de creación'] > '2024-01-01')].copy()), url=url_gs)
    
    url_av = 'https://script.google.com/macros/s/AKfycbwfOI88yhwtms6ofCfT_-j8pZDWZC4ph-hE90sliH2bpZUzCcwkK2WSaz9EQwxXhLw6Zg/exec'
    data_audiovisuales = df_to_array(df.loc[((df['Grupo'] == 'ESANLabs-Multimedia') | (df['Grupo'] == 'Enlaces Webinar')) & (df['Fecha de creación'] > '2024-01-01')].copy())
    Api('Post','TICKETS',data_audiovisuales, url=url_av)

    # PROCESAMIENTO LISTAS MICROSOFT
    df = df_tf.replace(pd.NaT, "-")

    df_media_2025 = df.loc[(df['Grupo'] == "ESANLabs-Multimedia") & (df['Fecha de creación'] > '2025-01-01') & (df['Asunto'].str.contains(r'[\(\)]+', na=False)) & (df['Estado'] != "Cerrado")]
    df_media_2025['Fecha de creación'] = df_media_2025['Fecha de creación'].astype(str)
    df_media_2025['Fecha de finalización'] = df_media_2025['Fecha de finalización'].astype(str)
    df_media_2025['Vencimiento antes de'] = df_media_2025['Vencimiento antes de'].astype(str)
    df_media_2025['Hora de inicio programada'] = df_media_2025['Hora de inicio programada'].astype(str)
    
    tickets_media_2025 = [dict(zip(df_media_2025.columns.values.tolist(), item)) for item in df_media_2025.values.tolist()]

    tickets_media_2025_ord = []
    
    SERVICIOS_DEFINIDOS = {
        'FOTO': ["foto", "fotografía", "fotografías", "fotos", "fot"],
        'VIDEO': ["video", "videos", "videografía", "grabación", "grabaciones", "grab", "fil", "fíl", "vid", "víd", "clip", "reel"],
        'VIDEORESUMEN': ["video resumen","videoresumen"],
        'STREAMING': ["streaming", "transmisión", "vivo"],
        'TESTIMONIO': ["testimonio", "testimonios", "testimo"],
        'EDICION': ["edición", "editar", "edicion", "reedición", "postproducción", "photoshop", "phot"],
        'ENTREVISTA': ["entrevista", "entrevistas"],
        'WEBINAR': ["zoom", "webinar", "videollamada", "sala", "virt", "link", "zoo", "zom"],
        'MATERIAL': ["material", "archivo", "envío de material"],
        'ANIMACION': ["animación", "animaciones", "animacion"],
        'COBERTURA': ["cobertura"],
        'PODCAST': ["podcast", "podc"],
        'OPENDAY': ["open"],
        'SPOT': ["spots","spot"]
    }

    if len(tickets_media_2025) > 0:
        for item in tickets_media_2025:
            try:
                found = re.findall(r'\((.*?)\)', item['Asunto'])
                if not found: continue
                    
                servicios = unidecode(found[0]).upper().split(",")

                for servicio in servicios:
                    servicio = servicio.strip()
                    item_format = {}
                    item_format['id'] = item['ID']
                    item_format['idMix'] = str(item['ID'])+ "-" + ''.join([palabra[0] for palabra in servicio.split()])
                    item_format['tituloSolicitud'] = re.sub(r'\(.*?\)', '', item['Asunto']).strip()
                    item_format['solicitante'] = item['Nombre del solicitante']
                    item_format['aSolicitante'] = item['Departamento']

                    if (item['Hora de inicio programada'].split(" ")[0]) == "-":
                        item_format['fEvento'] = ""
                    else:
                        try:
                            item_format['fEvento'] = datetime.strptime(item['Hora de inicio programada'].split(" ")[0], '%d/%m/%Y').strftime('%Y-%m-%d')
                        except:
                            item_format['fEvento'] = ""

                    servicio_limpio = servicio.strip().lower()
                    servicio_detectado = None
                    for categoria, sinonimos in SERVICIOS_DEFINIDOS.items():
                        if any(s in servicio_limpio for s in sinonimos):
                            servicio_detectado = categoria
                            break

                    item_format['servicio'] = servicio_detectado if servicio_detectado else servicio.strip().upper()
                    tickets_media_2025_ord.append(item_format)
            except Exception as e:
                print(f"Error procesando item {item['ID']}: {e}")

    # Power Automate
    url_pa = "https://prod-14.brazilsouth.logic.azure.com:443/workflows/7f928698304d46a9a3771f3936d51726/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=nL0VFc8HAwdLDplKCzWASz-GekhIRG73hGzVk3cLN3Y"
    
    print(f"Enviando {len(tickets_media_2025_ord)} registros a Power Automate...")
    for item in tickets_media_2025_ord:
        try:
            requests.post(url=url_pa, json=item)
        except Exception as e:
            print(f"Error enviando {item.get('idMix')}: {e}")

    # Limpieza final en Descargas
    if os.path.exists(full_path_file):
        os.remove(full_path_file)
        print("Limpieza final completada en la carpeta Descargas.")
else:
    print("No se pudo procesar porque el archivo no se encontró en Descargas.")

#automatizado by: Ely Bot
