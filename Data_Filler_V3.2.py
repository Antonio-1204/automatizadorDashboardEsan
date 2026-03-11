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

print("\n\n\n")
print("#" * 100)
print("Iniciando el proceso de llenado de datos")
print("Buscando el archivo de solicitudes en la carpeta de trabajo \n\n")
file_dwnload = True

if "Requestexport.XLSX" in os.listdir():
    print(f"Se encontró el archivo de solicitudes en la carpeta de trabajo - 'Requestexport.XLSX'")
    resp = input("Desea continuar con el archivo encontrado?  ([Y]yes/[N]no): ")
    print("#" * 100,end="\n\n")

    if (resp.lower() == 'y' or resp.lower() == 'yes'):
        print("Continuando con el archivo existente")
        file_dwnload = False
    else:
        file_dwnload = True
        # Eliminar el archivo existente
        os.remove("Requestexport.XLSX")
        print("Se ha eliminado el archivo existente")

if file_dwnload:
    print("No se encontró el archivo de solicitudes en la carpeta de trabajo")
    print("#" * 100,end="\n\n")

    # Credenciales
    url = "https://servicedesk.esan.edu.pe/"
    user = 'educacionadistancia'
    # pwd = input("Ingrese la clave de la cuenta para continuar: ")
    pwd = 'rthj6724'

    ###### Código de descarga de la información
    # Configura las opciones para el navegador
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--incognito")  # Activa el modo incógnito
    
    # Inicializa el controlador de Chrome usando webdriver_manager
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)

    try:
        # time.sleep(5) # No es estrictamente necesario si la carga es rápida, pero se puede dejar
        driver.get(url)
        # Redirección forzada
        url = 'https://servicedesk.esan.edu.pe/WOListView.do?viewID=3190&globalViewName=All_Requests'
        time.sleep(4)

        # Zona de Login
        search_user = driver.find_element(By.NAME, "j_username")
        search_user.send_keys(user)
        search_pwd = driver.find_element(By.NAME, "j_password")
        search_pwd.send_keys(pwd)
        search_pwd.send_keys(Keys.ENTER)

        time.sleep(3)

        # Navegación a la URL final
        driver.get(url)
        time.sleep(5)

        # --- CORRECCIÓN DEL ERROR DE SELENIUM AQUÍ ---
        # Usamos WebDriverWait para asegurar que los elementos sean clickeables
        wait = WebDriverWait(driver, 20) # Espera máxima de 20 segundos

        # 1. Clic en el menú de acciones
        print("Intentando abrir menú de acciones...")
        menu_actions = wait.until(EC.element_to_be_clickable((By.ID, "bulkactionsMenu")))
        menu_actions.click()
               
        # 2. Clic en la opción exportar (abre el modal)
        print("Esperando opción de exportar...")
        export_option = wait.until(EC.element_to_be_clickable((By.ID, "load_export_dialog")))
        export_option.click()
        
        # --- EL TRUCO ESTÁ AQUÍ ---
        # Le damos 1.5 segundos para que la animación del modal termine de cargar por completo.
        time.sleep(1.5) 
        
        # 3. Seleccionar el formato XLSX específicamente
        print("Seleccionando formato XLSX...")
        # Usamos presence_of_element_located (verifica que exista, no importa si está "animándose")
        radio_xlsx = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@name='export-format' and @value='XLSX']")))
        driver.execute_script("arguments[0].click();", radio_xlsx)

        time.sleep(0.5) # Breve pausa entre clics

        # 4. Hacer clic en el botón de Exportar
        print("Activando descarga...")
        btn_exportar = wait.until(EC.presence_of_element_located((By.ID, "export_requestlistview")))
        driver.execute_script("arguments[0].click();", btn_exportar)

        # Activación del gatillador de creación del archivo de descarga
        # También le ponemos espera explícita por seguridad
        btn_exportar = wait.until(EC.element_to_be_clickable((By.ID, "export_requestlistview")))
        btn_exportar.click()

        # Aceptar el guardado del archivo cuando aparezca la segunda ventana de chrome
        print("Esperando ventana de guardar...")
        while ("Requestexport.XLSX" not in os.listdir()):
            # Verifica si la ventana de descargas de Chrome está abierta
            # NOTA: Esto depende del idioma del SO. Si está en inglés sería 'Save As'
            windows = gw.getWindowsWithTitle('Guardar Como') 
            if not windows:
                windows = gw.getWindowsWithTitle('Save As') # Intento en inglés por si acaso
            
            if windows:
                print("La ventana de descargas de Chrome detectada.")
                pyautogui.press('enter')
                break
            time.sleep(1)  
            
        print(os.listdir())
        while("Requestexport.XLSX" not in os.listdir()):
            print("Esperando terminar la descarga")
            time.sleep(3)
        print("Descarga del archivo realizada correctamente")

    except Exception as e: # Captura genérica mejorada para ver el error real
        print(f'Ocurrió un error: {e}')
    finally:
        driver.quit()


# Zona de tratamiento de los datos
if "Requestexport.XLSX" in os.listdir(): # Validación extra por seguridad
    # Para la data de prueba
    df = pd.read_excel('Requestexport.XLSX', skiprows = 7, usecols = 'B:P')

    # Ajuste del nombre de columnas
    df.columns = ['ID', 'Asunto',  'Nombre del solicitante','Asignado a', 'Vencimiento antes de', 'Estado' ,'Fecha de creación', 'Prioridad', 'Grupo', 'Catálogo de servicios', 'Fecha', 'Fecha de finalización', 'H. Inicio', 'Departamento','Hora de inicio programada']

    # Reordenando las columnas
    df = df[['ID', 'Asunto', 'Nombre del solicitante', 'Asignado a', 'Vencimiento antes de', 'Estado', 'Fecha de creación', 'Prioridad', 'Grupo', 'Catálogo de servicios', 'Fecha', 'Fecha de finalización', 'H. Inicio', 'Departamento','Hora de inicio programada']]

    # Reordenando las filas
    df = df.sort_values(by='ID')

    # Ajuste del tiempo y asignación según corresponda
    # Para Fecha Creación
    df['Fecha de creación'] = pd.to_datetime(df['Fecha de creación'], format='%d/%m/%Y %I:%M %p')
    # Para fecha de finalización
    df.loc[df['Fecha de finalización'] == '-', 'Fecha de finalización'] = datetime.now().strftime('%d/%m/%Y %I:%M %p')
    df['Fecha de finalización'] = pd.to_datetime(df['Fecha de finalización'], format='%d/%m/%Y %I:%M %p')
    # Para la fecha de vencimiento asignado
    df.loc[df['Vencimiento antes de'] == '-', 'Vencimiento antes de'] = ''
    df['Vencimiento antes de'] = pd.to_datetime(df['Vencimiento antes de'], format='%d/%m/%Y %I:%M %p')
    # Creando un Dataframe con solo el tratamiento de fechas
    df_tf = df.copy()


    # Asignación de área a Martha, Gabriel y Andrea
    df.loc[df['Asignado a'] == 'Martha Medina Cloke', 'Grupo'] = 'ESANLabs-Diseño'
    df.loc[df['Asignado a'] == 'Gabriel Juarez Postigo', 'Grupo'] = 'ESANLabs - Ted'
    df.loc[df['Asignado a'] == 'Maria Jose Sanchez Stein', 'Grupo'] = 'ESANLabs - Ted'
    df.loc[df['Asignado a'].str.contains('Andrea', case = False), 'Grupo'] = 'ESANLabs - Ted'
    df.loc[df['Asunto'].str.contains('Reenviar este correo', case = False), 'Grupo'] = 'Enlaces Zoom'
    df.loc[df['Asunto'].str.contains('CAP -', case = False), 'Grupo'] = 'CAP'
    df.loc[df['Asunto'].str.contains('AESPI - ', case = False), 'Grupo'] = 'AESPI'
    df.loc[(df['Asunto'].str.contains('zoom', case = False)) & (df['Grupo'] == 'ESANLabs-Multimedia'), 'Grupo'] = 'Enlaces Webinar' 
    df.loc[df['Asunto'].str.contains('PRY', case = False), 'Grupo'] = "Esanlabs - Proyectos"
    df.loc[df['Asignado a'] == 'Yessica Renee Guardamino Vera', 'Grupo'] = 'ESANLabs - Ted'
    # Cambio de área de Juan Carlos según el tiempo
    df.loc[(df['Fecha de creación'] > '2024-08-01') & (df['Asignado a'].str.contains('Juan Carlos Leyva Carrasco', case = False)), 'Grupo'] = 'ESANLabs-Multimedia'
    df.loc[(df['Fecha de creación'] < '2024-08-01') & (df['Asignado a'].str.contains('Juan Carlos Leyva Carrasco', case = False)), 'Grupo'] = 'ESANLabs - Ted'

    #############################################################################################################################################################################
    # Zona de Funciones de carga y otros
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
            try:
                response_json = response.json()
                print("Status:", response_json['status'])
                return response.json()
            except requests.exceptions.JSONDecodeError:
                print("Failed to decode JSON. Response was:", response.text)
        except requests.exceptions.RequestException as e:
            print("Request failed:", e)

    def df_to_array(aux):
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

    #############################################################################################################################################################################
    # Carga de información de los tickets en general:
    url_gs = 'https://script.google.com/macros/s/AKfycby-d58Qp5lvoLHvBALoIVBIDudbEnOgDrGY-V9CNcs2xSo6lufl6TmtcwnhrVBjnM4-Zw/exec'
    Api('Post','GENERAL',df_to_array(df), url=url_gs)
    Api('Post','MULTIMEDIA',df_to_array(df.loc[((df['Grupo'] == 'ESANLabs-Multimedia') | (df['Grupo'] == 'Enlaces Webinar')) & (df['Fecha de creación'] > '2024-01-01')].copy()), url=url_gs)
    
    #############################################################################################################################################################################
    # Carga de la información para audiovisuales
    url_av = 'https://script.google.com/macros/s/AKfycbwfOI88yhwtms6ofCfT_-j8pZDWZC4ph-hE90sliH2bpZUzCcwkK2WSaz9EQwxXhLw6Zg/exec'
    data_audiovisuales = df_to_array(df.loc[((df['Grupo'] == 'ESANLabs-Multimedia') | (df['Grupo'] == 'Enlaces Webinar')) & (df['Fecha de creación'] > '2024-01-01')].copy())
    Api('Post','TICKETS',data_audiovisuales, url=url_av)

    ############################################################################################################################################################################
    # PROCESAMIENTO LISTAS MICROSOFT
    
    df = df_tf.replace(pd.NaT, "-")

    df_media_2025 = df.loc[(df['Grupo'] == "ESANLabs-Multimedia") & (df['Fecha de creación'] > '2025-01-01') & (df['Asunto'].str.contains(r'[\(\)]+')) & (df['Estado'] != "Cerrado")]
    df_media_2025['Fecha de creación'] = df_media_2025['Fecha de creación'].astype(str)
    df_media_2025['Fecha de finalización'] = df_media_2025['Fecha de finalización'].astype(str)
    df_media_2025['Vencimiento antes de'] = df_media_2025['Vencimiento antes de'].astype(str)
    df_media_2025['Hora de inicio programada'] = df_media_2025['Hora de inicio programada'].astype(str)
    
    tickets_media_2025 = [dict(zip(df_media_2025.columns.values.tolist(), item)) for item in df_media_2025.values.tolist()]

    tickets_media_2025_ord = [] # Aquí se guardarán los diccionarios ordenados
    
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

    if len(tickets_media_2025) > 0: # Evitar error si la lista está vacía
        for item in tickets_media_2025:
            try:
                # Verificar que el regex encontró algo
                found = re.findall(r'\((.*?)\)', item['Asunto'])
                if not found:
                    continue # Saltar si no hay paréntesis
                    
                servicios = unidecode(found[0]).upper().split(",")

                for servicio in servicios:
                    servicio = servicio.strip()
                    item_format = {}
                    item_format['id'] = item['ID']
                    item_format['idMix'] = str(item['ID'])+ "-" + ''.join([palabra[0] for palabra in servicio.split()])
                    item_format['tituloSolicitud'] = re.sub(r'\(.*?\)', '', item['Asunto']).strip() # Limpiamos espacios
                    item_format['solicitante'] = item['Nombre del solicitante']
                    item_format['aSolicitante'] = item['Departamento']

                    if (item['Hora de inicio programada'].split(" ")[0]) == "-":
                        item_format['fEvento'] = ""
                    else:
                        try:
                            item_format['fEvento'] = datetime.strptime(item['Hora de inicio programada'].split(" ")[0], '%d/%m/%Y').strftime('%Y-%m-%d')
                        except:
                            item_format['fEvento'] = "" # Fallback si falla la fecha

                    servicio_limpio = servicio.strip().lower()

                    servicio_detectado = None
                    for categoria, sinonimos in SERVICIOS_DEFINIDOS.items():
                        if any(s in servicio_limpio for s in sinonimos):
                            servicio_detectado = categoria
                            break

                    if servicio_detectado:
                        item_format['servicio'] = servicio_detectado
                    else:
                        item_format['servicio'] = servicio.strip().upper()

                    # --- CORRECCIÓN LÓGICA: AGREGAR A LA LISTA ---
                    tickets_media_2025_ord.append(item_format)
                    # ---------------------------------------------
            except Exception as e:
                print(f"Error procesando item {item['ID']}: {e}")

    # Carga de información segmentada a Power Automate:
    url_pa = "https://prod-14.brazilsouth.logic.azure.com:443/workflows/7f928698304d46a9a3771f3936d51726/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=nL0VFc8HAwdLDplKCzWASz-GekhIRG73hGzVk3cLN3Y"
    
    print(f"Enviando {len(tickets_media_2025_ord)} registros a Power Automate...")

    # Carga de la información diversificada a Microsoft Lists
    for item in tickets_media_2025_ord:
        try:
            r = requests.post(url=url_pa, json=item)
            # Opcional: imprimir status para verificar
            # print(f"Enviado {item['idMix']}: {r.status_code}")
        except Exception as e:
            print(f"Error enviando {item.get('idMix')}: {e}")

    if os.path.exists("Requestexport.XLSX"):
        os.remove("Requestexport.XLSX")
        print("Limpieza final completada.")
else:
    print("No se pudo procesar porque el archivo no se descargó.")

#automatizado by: Renzo Vargas