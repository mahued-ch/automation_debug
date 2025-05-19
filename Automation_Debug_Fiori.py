import sys
import os
# Asegura que la carpeta actual esté en el sys.path
sys.path.append(os.path.dirname(__file__))
import pyautogui
import time
import cv2
from datetime import datetime, timedelta
from pywinauto import Application
from Fiori_Reporte import get_reporte
from Automation_Variables import ruta, ruta_output, boton_flecha_atras, boton_ok_intro, grabar_lista_fichero, texto_con_tabuladores, grabar_fichero, boton_reemplazar, boton_permitir 
from Automation_Variables import boton_cancel, boton_auditoria1, boton_auditoria2, boton_reinicializar, log_debug, boton_seleccion_multiple, boton_arriba_abajo, boton_seleccionar_detalle
from Automation_Variables import boton_leer_auditoria, texto_grave, texto_nocritico_AU5, texto_nocritico_AUK, texto_nocritico_BU4, texto_critico, texto_con_tabuladores, ventana_noresultados
from Automation_Genera_Docto import Genera_Documento
from Automation_Variables import borrar_archivos_carpeta_con_prefijo, limpiar_pantalla
from Automation_Variables import configuracion, carga_config, region4, region1
from Automation_Functions import convert_to_excel_fiori, abrir_correo_outlook, espera_cambio_pantalla, captura_pantalla, busca_en_pantalla
from Automation_SAP import start_sap_logon, login_to_sap, close_window, close_sap_logon
import Automation_Variables

# ###--------------------------------------------------------------------------------
# def busca_en_pantalla(imagen, intervalo, reintentos, confidence=0.8): 
#     location = None
#     conta = 0
#     while (location == None) and (conta <= reintentos):
#         try:
#             conta += 1 
#             time.sleep(intervalo)
#             location = pyautogui.locateCenterOnScreen(imagen, confidence=confidence)
#         except pyautogui.ImageNotFoundException:
#             print(f"Imagen no encontrada en intento {conta}.")
#             continue
#         except Exception as e:
#             print(f"Error inesperado: {e}")
#             break
#     return location

###--------------------------------------------------------------------------------
def buscar_todas_ocurrencias(imagen, confidence=0.8):
#    Busca todas las posiciones de una imagen en la pantalla.
    posiciones = None
    while (posiciones == None) :
        try:
            posiciones = list(pyautogui.locateAllOnScreen(imagen, confidence=confidence))
        except Exception:
            pass
            break
    return posiciones

###--------------------------------------------------------------------------------
def hazclick_en_evento(tipo_evento):
  for lc_posicion in tipo_evento:
      # Calcular la posición de la casilla de selección
      x_seleccion = lc_posicion.left + offset_x
      y_seleccion = lc_posicion.top + offset_y
    # Hacer clic en la casilla
      pyautogui.click(x_seleccion, y_seleccion)
      time.sleep(1)  # Pequeña pausa entre clics
  return True

###--------------------------------------------------------------------------------

fecha_actual = datetime.now()
# Obtener el lunes anterior menos 1 día
fecha_actual = fecha_actual - timedelta(days=fecha_actual.weekday() + 1)
# Calcular la fecha inicial restando 6 días
fecha_anterior = fecha_actual - timedelta(days=6)

fecha_amd = fecha_actual.strftime("%Y%m%d")

fecha_1 = fecha_anterior.strftime("%d.%m.%Y")
fecha_2 = fecha_actual.strftime("%d.%m.%Y")
hora_1  = '00:00:00'
hora_2  = '23:59:59'

limpiar_pantalla()

configuracion = carga_config()

Automation_Variables.running_test = int(configuracion["config"]["testing"]) 

start_sap_logon(configuracion)
login_to_sap(configuracion["sap"]["system_fiori"], configuracion["fiori"]["usuario"], configuracion["fiori"]["password"])

prefijo = "Fiori_criticos"
contador = 1

# Borramos todos los archivos de la ruta de salida
if Automation_Variables.running_test >= 0 :
  borrar_archivos_carpeta_con_prefijo(ruta_output, prefijo)

  # buscamos último archivo de captura generado
  while os.path.exists(os.path.join(ruta_output, f"{prefijo}_{fecha_amd}_{contador}.png")):
      contador += 1

  # Esperar 3 segundos para permitirte enfocar el formulario
#  time.sleep(3)

  Automation_Variables.file_output = 'dentro_de_sm20n'

  # Presionar Shift-F5
  print("Proporcionando fechas...")
  pyautogui.keyDown('shift')  # Mantiene presionada la tecla Shift
  pyautogui.press('tab')       # Presiona tab
  pyautogui.keyUp('shift')    # Suelta la tecla Shift

  pyautogui.typewrite(fecha_1)
  pyautogui.press('tab')       # Presiona tab
  pyautogui.typewrite(hora_1)
  pyautogui.press('tab')       # Presiona tab
  pyautogui.typewrite(fecha_2)
  pyautogui.press('tab')       # Presiona tab
  pyautogui.typewrite(hora_2)
  pyautogui.press('tab')       # Presiona tab

  # Reinicializamos lista
  print("Presionamos botón reinicializar.")
  location = busca_en_pantalla([boton_reinicializar], 1, 10) 
  if location != None:
    pyautogui.click(location)

  print("Presionamos botón seleccionar detalle.")
  location = busca_en_pantalla([boton_seleccionar_detalle], 1, 10) 
  if location != None:
    pyautogui.click(location)

  print("Presionamos botón cancelar.")
  location = busca_en_pantalla([boton_cancel], 1, 10) 
  if location != None:
    pyautogui.click(location)

  # Reinicializamos lista
  print("Presionamos botón reinicializar.")
  location = busca_en_pantalla([boton_reinicializar], 1, 10) 
  if location != None:
    pyautogui.click(location)

  print("Presionamos botón seleccionar detalle.")
  location = busca_en_pantalla([boton_seleccionar_detalle], 1, 10) 
  if location != None:
    pyautogui.click(location)

  for i in range(15):
    print(f"Mostramos la página {i}.")

    hubo_click = False
    time.sleep(2)  # Pequeña pausa entre clics
  # Si es la última página me regreso 2 clicks para tomar en cuenta el último crítico
#    if i == 6:
#      location = busca_en_pantalla([boton_arriba_abajo], 1, 10) 
#      location = (location[0], location[1] + 15)
#      pyautogui.click(location)
#      pyautogui.click(location)
#      time.sleep(2)  # Pequeña pausa entre clics

    eventos_criticos = None
    if i <= 6:
      eventos_criticos = buscar_todas_ocurrencias(texto_critico, confidence=0.9)

    if i == 6 and eventos_criticos:
      eventos_criticos = [eventos_criticos[0]]  # Solo conservar la primera
    
    offset_x = 118  # Ajusta la distancia al botón "Selección" (en X)
    offset_y = 7    # Ajusta si el botón está más arriba o abajo

    if eventos_criticos:
        print(f"Se encontraron {len(eventos_criticos)} eventos críticos.")

        hubo_click = hazclick_en_evento(eventos_criticos)
        
    else:
        print("No se encontraron eventos críticos.")

  # buscamos eventos no críticos 
    eventos_nocriticos = None
    if i == 0: # Se encuentra en la primera página
      print("Buscamos BU4")
      eventos_nocriticos = buscar_todas_ocurrencias(texto_nocritico_BU4, confidence=0.90)
    elif i == 10:  
      print("Buscamos AU5")
      eventos_nocriticos = buscar_todas_ocurrencias(texto_nocritico_AU5, confidence=0.90)
    elif i == 12: 
      print("Buscamos AUK")
      eventos_nocriticos = buscar_todas_ocurrencias(texto_nocritico_AUK, confidence=0.90)

    if eventos_nocriticos:
        print(f"Se encontraron {len(eventos_nocriticos)} eventos nocríticos.")

        hubo_click = hazclick_en_evento(eventos_nocriticos)

    else:
      print("No se encontraron eventos no críticos.")

    if hubo_click:
      time.sleep(1)
      screenshot = pyautogui.screenshot()
      print(f"Grabando archivo {prefijo}_{fecha_amd}_{contador}.png")
      screenshot.save(os.path.join(ruta_output, f"{prefijo}_{fecha_amd}_{contador}.png"))    
      contador += 1

  # tomo captura de cada una de las pantallas de valores seleccionados
    captura_inicial = captura_pantalla((region4))    
    pyautogui.press('pagedown') 
    cambio_detectado = espera_cambio_pantalla(2, 10, (region4), captura_inicial)
    if not cambio_detectado:
      print('automation_debug_fiori - 1 - No se encontró cambio de pantalla')
      sys.exit()

  location = busca_en_pantalla([boton_ok_intro], 1, 10) 
  if location != None:
    pyautogui.click(location)
    time.sleep(1)
    screenshot = pyautogui.screenshot()
    print(f"Grabando archivo {prefijo}_{fecha_amd}_{contador}.png")
    screenshot.save(os.path.join(ruta_output, f"{prefijo}_{fecha_amd}_{contador}.png"))    
    contador += 1

  # Presionamos leer auditoria
  # hacemos el intento de ejecutar el botón durante 5 segundos seguidos, para que no se pierda
  # o si la pantalla cambia antes de terminar esos 5 segundos se sale
  print("Presionamos botón leer auditoria...")
  captura_inicial = captura_pantalla((region1))    
  for _ in range(5):
      pyautogui.press('F8')       # Presiona F8
      cambio_detectado = espera_cambio_pantalla(1, 1, (region1), captura_inicial)
      if cambio_detectado == True:
        break
      time.sleep(1)  # Espera 1 segundo entre cada pulsación    

    
#  location = busca_en_pantalla([boton_leer_auditoria], 1, 5) 
#  if location != None:
#    pyautogui.click(location)

  # Buscamos si no tuvo resultados
  time.sleep(5)
  print("Buscamos si no tuvo resultados...")

  while True:
    locationnr = busca_en_pantalla([ventana_noresultados], 2, 1)
    # Buscamos el cambio de pantalla con el botón de auditoría
    locationau = busca_en_pantalla([boton_auditoria2], 5, 1) 
    if (locationnr == None) and (locationau == None):
  # si todavía no aparece ninguna otra pantalla vamos por la siguiente ejecución    
      time.sleep(5)
      continue
    elif (locationnr != None):
  # si se encontró el botón de que no hubo datos presiona el botón intro y termina
      location = busca_en_pantalla([boton_ok_intro], 1, 5) 
      break
    elif (locationau != None):
  # si ya encontró el botón de auditoría continua con la ejecución del proceso
        get_reporte(prefijo, fecha_amd)
        time.sleep(1)
    # Grabamos la pantalla
        screenshot = pyautogui.screenshot()
        print(f"Grabando archivo {prefijo}_{fecha_amd}_{contador}.png")
        screenshot.save(os.path.join(ruta_output, f"{prefijo}_{fecha_amd}_{contador}.png"))    
        contador += 1
        captura_inicial = captura_pantalla((region4))    
        pyautogui.press('f3')       # Presiona F3
        cambio_detectado = espera_cambio_pantalla(2, 10, (region4), captura_inicial)
        break
      
    
  # Generamos el documento
  Genera_Documento(f"{prefijo}_{fecha_amd}.docx", prefijo, 'LOG de Debug para clase de eventos “Crítico” y “No crítico”')

print("Conversión de archivo obtenido a .xlsx ...")
resultado = convert_to_excel_fiori(os.path.join(ruta_output, f"{prefijo}_{fecha_amd}.xls"))

if "archivo_salida" in resultado:
  archivos = [resultado["archivo_salida"]]  # Lista de archivos a adjuntar    
else:
  archivos = []
    
prefijo2 = f"{prefijo}_{fecha_amd}"
ruta = ruta_output

print("Esperamos 5 segundos ...")
time.sleep(5)

try:    
  print("Intentando abrir correo de eventos críticos y no críticos ...")
  abrir_correo_outlook(configuracion, configuracion["email"]["subject_fiori_1"], resultado, fecha_1, fecha_2, configuracion["email"]["recipient"], archivos, ruta, prefijo2, ['.png', '.docx'])    
except Exception as e:
  os.startfile(ruta)  
  print(f"Error al iniciar Outlook: {e}")

########################################################################
# Segunda parte, vamos por los graves 
########################################################################
prefijo = "Fiori_graves"

# Borramos todos los archivos de la ruta de salida
if Automation_Variables.running_test >= 0:
  borrar_archivos_carpeta_con_prefijo(ruta_output, prefijo)

  contador = 1
  # buscamos último archivo de captura generado
  while os.path.exists(os.path.join(ruta_output, f"{prefijo}_{fecha_amd}_{contador}.png")):
      contador += 1

  # Reinicializamos lista
  print("Presionamos botón reinicializar.")
  location = busca_en_pantalla([boton_reinicializar], 1, 10) 
  if location != None:
    pyautogui.click(location)

  print("Presionamos botón seleccionar detalle.")
  location = busca_en_pantalla([boton_seleccionar_detalle], 1, 10) 
  if location != None:
    pyautogui.click(location)

  Automation_Variables.file_output = 'seleccion_graves'

  for i in range(7):
    print(f"Mostramos la página {i}.")

    hubo_click = False
    time.sleep(2)  # Pequeña pausa entre clics
  # Si es la última página me regreso 2 clicks para tomar en cuenta el último crítico

    eventos_graves = None
    eventos_graves = buscar_todas_ocurrencias(texto_grave, confidence=0.9)

    if i == 6 and eventos_graves:
      eventos_graves = [eventos_graves[0]]  # Solo conservar la primera

    offset_x = 118  # Ajusta la distancia al botón "Selección" (en X)
    offset_y = 7    # Ajusta si el botón está más arriba o abajo

    if eventos_graves:
        print(f"Se encontraron {len(eventos_graves)} eventos graves.")

        hubo_click = hazclick_en_evento(eventos_graves)
        
    else:
        print("No se encontraron eventos críticos.")

    if hubo_click:
      time.sleep(1)
      screenshot = pyautogui.screenshot()
      screenshot.save(os.path.join(ruta_output, f"{prefijo}_{fecha_amd}_{contador}.png"))    

      contador += 1

  # tomo captura de cada una de las pantallas de valores seleccionados
    captura_inicial = captura_pantalla((region4))    
    pyautogui.press('pagedown') 
    cambio_detectado = espera_cambio_pantalla(2, 10, (region4), captura_inicial)
    if not cambio_detectado:
      print('automation_debug_fiori - 2 - No se encontró cambio de pantalla')
      sys.exit()

  location = busca_en_pantalla([boton_ok_intro], 1, 10) 
  if location != None:
    pyautogui.click(location)
    time.sleep(1)
    screenshot = pyautogui.screenshot()
    screenshot.save(os.path.join(ruta_output, f"{prefijo}_{fecha_amd}_{contador}.png"))    
    contador += 1

  # Presionamos leer auditoria
  print("Presionamos botón leer auditoria...")
  pyautogui.press('F8')       # Presiona F8
#  location = busca_en_pantalla([boton_leer_auditoria], 1, 5) 
#  if location != None:
#    pyautogui.click(location)

  # Buscamos si no tuvo resultados
  time.sleep(5)
  print("Buscamos si no tuvo resultados...")

  while True:
    locationnr = busca_en_pantalla([ventana_noresultados], 2, 1)
    # Buscamos el cambio de pantalla con el botón de auditoría
    locationau = busca_en_pantalla([boton_auditoria2], 5, 1) 
    if (locationnr == None) and (locationau == None):
  # si todavía no aparece ninguna otra pantalla vamos por la siguiente ejecución    
      time.sleep(5)
      continue
    elif (locationnr != None):
  # grabamos la pantalla    
      contador += 1
      time.sleep(1)
      screenshot = pyautogui.screenshot()
      print(f"Grabando archivo {prefijo}_{fecha_amd}_{contador}.png")
      screenshot.save(os.path.join(ruta_output, f"{prefijo}_{fecha_amd}_{contador}.png"))    
      contador += 1
  # si se encontró el botón de que no hubo datos presiona el botón intro y termina
      location = busca_en_pantalla([boton_ok_intro], 1, 5) 
      pyautogui.click(location)
      break
    elif (locationau != None):
  # si ya encontró el botón de auditoría continua con la ejecución del proceso
        get_reporte(prefijo, fecha_amd)
        time.sleep(1)
    # Grabamos la pantalla
        screenshot = pyautogui.screenshot()
        print(f"Grabando archivo {prefijo}_{fecha_amd}_{contador}.png")
        screenshot.save(os.path.join(ruta_output, f"{prefijo}_{fecha_amd}_{contador}.png"))    
        contador += 1
        pyautogui.press('f3')       # Presiona F3
        break
        

  # Generamos el documento
  Genera_Documento(f"{prefijo}_{fecha_amd}.docx", prefijo, 'LOG de Debug para clase de eventos “Grave”')

print("Conversión de archivo obtenido a .xlsx ...")
#resultado = convert_to_excel_fiori(f"{prefijo}_{fecha_amd}.xls")
resultado = convert_to_excel_fiori(os.path.join(ruta_output, f"{prefijo}_{fecha_amd}.xls"))

if "archivo_salida" in resultado:
  archivos = [resultado["archivo_salida"]]  # Lista de archivos a adjuntar    
else:
  archivos = []

close_window()
close_sap_logon()
      
prefijo2 = f"{prefijo}_{fecha_amd}"
ruta = ruta_output
try:    
  print("Intentando abrir correo de eventos graves ...")
  abrir_correo_outlook(configuracion, configuracion["email"]["subject_fiori_2"], resultado, fecha_1, fecha_2, configuracion["email"]["recipient"], archivos, ruta, prefijo2, ['.png', '.docx'])    
except Exception as e:
  os.startfile(ruta)  
  print(f"Error al iniciar Outlook: {e}")

########################################################################
# Fin del programa
########################################################################

print("Fin del Programa")

sys.exit()  


