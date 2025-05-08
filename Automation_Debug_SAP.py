import sys
import os
# Asegura que la carpeta actual esté en el sys.path
sys.path.append(os.path.dirname(__file__))
import pyautogui
import time
from PIL import ImageChops
from datetime import datetime, timedelta
from Automation_Variables import ruta, ruta_output, boton_flecha_atras, boton_ok_intro, grabar_lista_fichero, texto_con_tabuladores, grabar_fichero, boton_reemplazar, boton_permitir 
from Automation_Variables import ruta_imagenes, boton_auditoria1, boton_auditoria2, boton_reinicializar, log_debug, boton_seleccion_multiple, boton_arriba_abajo, boton_seleccionar_detalle
from Automation_Variables import boton_leer_auditoria, texto_grave, texto_nocritico_AU5, texto_nocritico_AUK, texto_nocritico_BU4, texto_critico, texto_con_tabuladores, ventana_noresultados
from Automation_Genera_Docto import Genera_Documento
from Automation_Variables import borrar_archivos_carpeta_con_prefijo, limpiar_pantalla, running_test 
from Automation_Variables import configuracion, carga_config
from Automation_Functions import convert_to_excel, abrir_correo_outlook, captura_pantalla, espera_cambio_pantalla, copiar_imagen_al_clipboard
from Automation_SAP import start_sap_logon, login_to_sap, close_window, close_sap_logon
from Automation_Variables import region1, region3

def pantallas_iguales(img1, img2):
    return ImageChops.difference(img1, img2).getbbox() is None
  
def busca_en_pantalla(imagen, intervalo, reintentos, confidence=0.8): 
    location = None
    conta = 0
    while (location == None) and (conta <= reintentos):
        try:
            conta += 1 
            time.sleep(intervalo)
            location = pyautogui.locateCenterOnScreen(imagen, confidence=confidence)
        except pyautogui.ImageNotFoundException:
            print(f"Imagen no encontrada en intento {conta}.")
            continue
        except Exception as e:
            print(f"Error inesperado: {e}")
            break
    return location

limpiar_pantalla()

configuracion = carga_config()

start_sap_logon(configuracion)
login_to_sap(configuracion["sap"]["system_s4p"], configuracion["s4p"]["usuario"], configuracion["s4p"]["password"])

# calculamos rango de fechas
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

prefijo = "s4p"

contador = 1

if running_test == 0:
  # Borramos todos los archivos de la ruta de salida
  borrar_archivos_carpeta_con_prefijo(ruta_output, prefijo)

  # buscamos último arcihvo de captura generado
  while os.path.exists(os.path.join(ruta_output, f"{prefijo}_{fecha_amd}_{contador}.png")):
      contador += 1

  # Esperar 3 segundos para permitirte enfocar el formulario
  time.sleep(3)

  # Presionar Shift-F5
  captura_inicial = captura_pantalla(region1)    
  
  print("Buscamos variante...")
  pyautogui.keyDown('shift')  # Mantiene presionada la tecla Shift
  pyautogui.press('f5')       # Presiona F5
  pyautogui.keyUp('shift')    # Suelta la tecla Shift
  cambio_detectado = espera_cambio_pantalla(2, 10, (region1), captura_inicial)
  if not cambio_detectado:
    print('automation_debug_sap - 1 - No se encontró ventana de SAP')
    sys.exit()

  # selecciona la variante
  print("Seleccionamos variante...")
  time.sleep(1)
  location = busca_en_pantalla(log_debug, 1, 10, 0.7) 
  if location != None:
    captura_inicial = captura_pantalla((region1))    
    pyautogui.doubleClick(location)
    cambio_detectado = espera_cambio_pantalla(2, 10, (region1), captura_inicial)
    if not cambio_detectado:
      print('automation_debug_sap - 2 - No se encontró ventana de SAP')
      sys.exit()
  else:
      print("No se encontró 'LOGDEBUG', salimos del programa...")
      sys.exit()  

  pyautogui.typewrite(fecha_1)
  pyautogui.press('tab')       # Presiona tab
  pyautogui.typewrite(hora_1)
  pyautogui.press('tab')       # Presiona tab
  pyautogui.typewrite(fecha_2)
  pyautogui.press('tab')       # Presiona tab
  pyautogui.typewrite(hora_2)
  pyautogui.press('tab')       # Presiona tab

  ###
  # Capturamos la primera pantalla, donde está la fecha
  screenshot = pyautogui.screenshot()
  print(f"Grabando archivo {prefijo}_{fecha_amd}_{contador}.png")
  screenshot.save(os.path.join(ruta_output, f"{prefijo}_{fecha_amd}_{contador}.png"))    
  contador += 1

  print("Botón selección múltiple...")
  # boton para mostrar selección
  location = busca_en_pantalla(boton_seleccion_multiple, 1, 10) 
  if location != None:
    captura_inicial = captura_pantalla((region1))    
    pyautogui.click(location)
    cambio_detectado = espera_cambio_pantalla(2, 10, (region1), captura_inicial)
    if not cambio_detectado:
      print('automation_debug_sap - 3 - No se encontró ventana de SAP')
      sys.exit()

  # tomo captura de cada una de las pantallas de valores seleccionados
#  location = busca_en_pantalla(boton_arriba_abajo, 1, 10) 
#  location = (location[0], location[1] - 25)

  print("Capturamos pantallas...")
  for i in range(6):
#      time.sleep(1)
      screenshot = pyautogui.screenshot()
      print(f"Grabando archivo {prefijo}_{fecha_amd}_{contador}.png")
      screenshot.save(os.path.join(ruta_output, f"{prefijo}_{fecha_amd}_{contador}.png"))    
      pyautogui.click(location)
      captura_inicial = captura_pantalla((region1))    
      pyautogui.press('pagedown')       # Presiona F8
      cambio_detectado = espera_cambio_pantalla(2, 10, (region1), captura_inicial)
      if not cambio_detectado:
        print('automation_debug_sap - 4 - No se encontró ventana de SAP')
        sys.exit()
      contador += 1

  # salgo de la pantalla de selección
#  time.sleep(1)
  captura_inicial = captura_pantalla((region1))    
  pyautogui.press('f8')       # Presiona F8
  cambio_detectado = espera_cambio_pantalla(2, 10, (region1), captura_inicial)
  if not cambio_detectado:
    print('automation_debug_sap - 5 - No se encontró ventana de SAP')
    sys.exit()

  print("Ejecuta consulta...")
  # ejecuto la consulta
#  time.sleep(1)
  pyautogui.press('f8')       # Presiona F8
  captura_inicial = captura_pantalla((region1)) 
  copiar_imagen_al_clipboard(captura_inicial)
  time.sleep(3)
  cambio_detectado = espera_cambio_pantalla(5, 120, (region1), captura_inicial)
  if not cambio_detectado:
    print('automation_debug_sap - 6 - Consulta de auditoría no se ejecutó')
    sys.exit()

#  location = busca_en_pantalla(boton_auditoria1, 5, 120, 0.7) # esperamos hasta 10 minutos
#  if location == None:
#      print("No se encontró 'BOTON_AUDITORIA1', salimos del programa...")
#      sys.exit()  

  print("Selección de columnas...")
  # Presionar Ctrl-F8
  captura_inicial = captura_pantalla((region1))    
  pyautogui.keyDown('ctrl')  # Mantiene presionada la tecla Control
  pyautogui.press('f8')       # Presiona F8
  pyautogui.keyUp('ctrl')    # Suelta la tecla Shift
  cambio_detectado = espera_cambio_pantalla(2, 10, (region1), captura_inicial)
  if not cambio_detectado:
    print('automation_debug_sap - 7 - No se encontró ventana de SAP')
    sys.exit()

#  time.sleep(2)

  # Adiciona todas las columnas
  print("Adicionando columnas...")
  location = busca_en_pantalla(boton_flecha_atras, 1, 10) 
  location = (location[0], location[1] + 15)
  for i in range(10):
#      captura_inicial = captura_pantalla((6, 170, 850, 830))    
      captura_inicial = captura_pantalla((region3))    
      pyautogui.click(location)
      cambio_detectado = espera_cambio_pantalla(1, 10, (region3), captura_inicial)
      if not cambio_detectado:
        print('automation_debug_sap - 8 - No se encontró ventana de SAP')
#      time.sleep(1)

# presionamos OK dentro de la ventana de modificar layout
  location = busca_en_pantalla(boton_ok_intro, 1, 10) 
  if location != None:
    captura_inicial = captura_pantalla((region1))    
    pyautogui.click(location)
    cambio_detectado = espera_cambio_pantalla(1, 10, (region1), captura_inicial)
    if not cambio_detectado:
      print('automation_debug_sap - 9.1 - No se encontró ventana de SAP')
#    time.sleep(1)

  # Guardamos la pantalla.
  # le damos mas tiempo a que desaparezca el ok
  print("Esperando pantalla...")
  time.sleep(5)
  screenshot = pyautogui.screenshot()
  print(f"Grabando archivo {prefijo}_{fecha_amd}_{contador}.png")
  screenshot.save(os.path.join(ruta_output, f"{prefijo}_{fecha_amd}_{contador}.png"))    

  contador += 1
      
  # Presionar Ctrl-Shift-F9
  captura_inicial = captura_pantalla((region1))    
  pyautogui.keyDown('ctrl')  
  pyautogui.keyDown('shift') 
  pyautogui.press('f9')      
  pyautogui.keyUp('shift')   
  pyautogui.keyUp('ctrl')  
  cambio_detectado = espera_cambio_pantalla(5, 12, (region1), captura_inicial)
  if not cambio_detectado:
    print('automation_debug_sap - 9.2 - No se encontró ventana de SAP')

  print("Esperando Grabar lista fichero...")
  location = busca_en_pantalla(grabar_lista_fichero, 5, 12) 

#  ahora buscamos presionar el texto con tabuladores
  print("Presionamos texto con tabuladores...")
  location = pyautogui.locateCenterOnScreen(texto_con_tabuladores, confidence=0.8)
  pyautogui.click(location)
  time.sleep(1)

  # hago click en boton ok
  print("Presionamos botón ok...")
  location = busca_en_pantalla(boton_ok_intro, 1, 5) 
  if location != None:
    captura_inicial = captura_pantalla((region1))    
    pyautogui.click(location)
    cambio_detectado = espera_cambio_pantalla(2, 20, (region1), captura_inicial)
    if not cambio_detectado:
      print('automation_debug_sap - 10.1 - No se encontró ventana de SAP')

  # si ya está esperando el cambio de pantalla ya no sería necesario esperar grabar fichero...      
#  captura_inicial = captura_pantalla((region1))    
#  print("Esperando Grabar fichero...")
#  cambio_detectado = espera_cambio_pantalla(3, 60, (region1), captura_inicial)
#  if not cambio_detectado:
#    print('automation_debug_sap - 10.2 - No se encontró ventana de SAP')
      
  print("Ingresamos nombre del archivo...")
  pyautogui.typewrite(f"{prefijo}_{fecha_amd}.xls")

  # Presionar Shift-Tab
  pyautogui.keyDown('shift') 
  pyautogui.press('tab')      
  pyautogui.keyUp('shift')   

  # Y teclea la ruta del archivo de salida
  print("Ingresamos ruta del archivo...")
  pyautogui.typewrite(ruta_output)

  # Esperando botón reemplazar
  print("Esperando Botón Reemplazar...")
  location = busca_en_pantalla(boton_reemplazar, 2, 20) 
  if location != None:
    pyautogui.click(location)

  # Esperando botón permitir
  print("Esperando Botón Permitir...")
  location = busca_en_pantalla(boton_permitir, 2, 20) 
  if location != None:
    pyautogui.click(location)

  # Generamos el documento
  Genera_Documento(f"{prefijo}_{fecha_amd}.docx", prefijo, 'LOG de Debug para clase de eventos “Crítico”, “Grave” y “No crítico”')

#resultado = convert_to_excel(f"{prefijo}_{fecha_amd}.xls")
print("Conversión de archivo obtenido a .xlsx ...")
resultado = convert_to_excel(os.path.join(ruta_output, f"{prefijo}_{fecha_amd}.xls"))

if "archivo_salida" in resultado:
  archivos = [resultado["archivo_salida"]]  # Lista de archivos a adjuntar    
else:
  archivos = []
        
prefijo2 = f"{prefijo}_{fecha_amd}"
ruta = ruta_output

close_window()
close_sap_logon()

try:    
  print("Intentando abrir correo de eventos críticos, graves y no críticos ...")
#  abrir_correo_outlook(resultado, fecha_1, fecha_2, "mahued@chedraui.com.mx", archivos, ruta, prefijo2, ['.png', '.docx'])    
  abrir_correo_outlook(configuracion, configuracion["email"]["subject_s4p"], resultado, fecha_1, fecha_2, configuracion["email"]["recipient"], archivos, ruta, prefijo2, ['.png', '.docx'])    
  
except Exception as e:
  os.startfile(ruta)  
  print(f"Error al iniciar Outlook: {e}")

sys.exit()

