import pyautogui
import time
import os
import cv2
import sys
from datetime import datetime, timedelta
from Automation_Functions import busca_en_pantalla, buscar_todas_ocurrencias, limpiar_pantalla, espera_cambio_pantalla, captura_pantalla, buscar_centro_en_pantalla
from Automation_Variables import ruta, ruta_output, boton_flecha_atras, boton_ok_intro, grabar_lista_fichero, texto_con_tabuladores, grabar_fichero, boton_reemplazar, boton_permitir, region1, region3
import Automation_Variables

def get_reporte(lcprefijo, lcfecha_amd):
#    Ejecuta el paso de generación de reportes en el sistema Fiori
#    Este va a servir para los críticos y no críticos y para los graves.
  print("Selección de columnas...")
  # Presionar Ctrl-F8
  Automation_Variables.file_output = 'llamando_seleccion_columnas' ; 
  captura_inicial = captura_pantalla((region1))    
  pyautogui.keyDown('ctrl')  # Mantiene presionada la tecla Control
  pyautogui.press('f8')       # Presiona F8
  pyautogui.keyUp('ctrl')    # Suelta la tecla Shift
  cambio_detectado = espera_cambio_pantalla(2, 10, (region1), captura_inicial)
  if not cambio_detectado:
    print('fiori_reporte - 1 - No se encontró ventana de SAP')
    sys.exit()
  
#  time.sleep(2)

  # Adiciona todas las columnas
  print("Adicionando columnas...")
  Automation_Variables.file_output = 'adicionando_columnas' ; 
  location = busca_en_pantalla(boton_flecha_atras, 2, 10) 
  location = (location[0], location[1] + 15)
  for i in range(8):
      captura_inicial = captura_pantalla((region3))    
      pyautogui.click(location)
      cambio_detectado = espera_cambio_pantalla(2, 10, (region3), captura_inicial)
      if not cambio_detectado:
        print('fiori_reporte - 2 - No se encontró ventana de SAP')
#      time.sleep(1)

  Automation_Variables.file_output = 'boton_ok' ; 
  captura_inicial = captura_pantalla((region1))    
  location = busca_en_pantalla(boton_ok_intro, 1, 10) 
  if location != None:
    pyautogui.click(location)
    cambio_detectado = espera_cambio_pantalla(1, 10, (region1), captura_inicial)
    if not cambio_detectado:
      print('fiori_reporte - 3 - No se encontró cambio de pantalla')
      sys.exit()
#    time.sleep(1)

  max_intentos = 3
  location = None

  for intento in range(1, max_intentos + 1):
      print(f"\n🔁 Intento {intento} de {max_intentos}")

      # Captura inicial de pantalla en región 1
      Automation_Variables.file_output = 'grabar_lista' ; 
      captura_inicial = captura_pantalla(region1)    

      print("Presiono Ctrl-Shift-F9 ...")
      pyautogui.keyDown('ctrl')  
      pyautogui.keyDown('shift') 
      pyautogui.press('f9')      
      pyautogui.keyUp('shift')   
      pyautogui.keyUp('ctrl')  

      print("Esperando Grabar lista fichero...")
      cambio_detectado = espera_cambio_pantalla(1, 10, region1, captura_inicial)

      if not cambio_detectado:
          print('❌ fiori_reporte - 4 - No se encontró cambio de pantalla')
          continue  # Reintenta si no hubo cambio

      # Buscar en pantalla
      location = busca_en_pantalla(grabar_lista_fichero, 5, 12)

      if location is not None:
          print("✅ Ubicación encontrada.")
          break  # Éxito, salimos del bucle
      else:
          print("⚠️ No se encontró la ubicación. Reintentando...")

  if location is None:
      print("❌ Error: No se pudo encontrar la ubicación después de 3 intentos.")
      sys.exit()
 
  location = buscar_centro_en_pantalla(texto_con_tabuladores, confidence=0.8)  
# aquí no debería de tronar sino encuentra la imagen, algo debemos de hacer  
  pyautogui.click(location)
  time.sleep(1)

  # hago click en boton ok
  location = busca_en_pantalla(boton_ok_intro, 1, 5) 
  if location != None:
    pyautogui.click(location)
    time.sleep(1)

  print("Esperando Grabar fichero...")

  location = busca_en_pantalla(grabar_fichero, 3, 60) 

  time.sleep(1)
  pyautogui.typewrite(f"{lcprefijo}_{lcfecha_amd}.xls")

  time.sleep(2)
  # Presionar Shift-tab para regresal al campo de la ruta
  pyautogui.keyDown('shift') 
  pyautogui.press('tab')      
  pyautogui.keyUp('shift')   

  # Presionar Ctrl-A
  pyautogui.keyDown('ctrl') 
  pyautogui.press('a')      
  pyautogui.keyUp('ctrl')   

  # Presionar Delete
  pyautogui.keyDown('delete') 

  # Y teclea la ruta del archivo de salida
  pyautogui.typewrite(ruta_output)

  # Esperando botón reemplazar
  print("Esperando Botón Reemplazar...")
  location = busca_en_pantalla(boton_reemplazar, 1, 20) 
  if location != None:
    pyautogui.click(location)

  # Esperando botón reemplazar
  print("Esperando Botón Permitir...")
  Automation_Variables.file_output = 'boton_permitir' ; 
  captura_inicial = captura_pantalla((region1))    
  location = busca_en_pantalla(boton_permitir, 1, 20) 
  if location != None:
    pyautogui.click(location)
    cambio_detectado = espera_cambio_pantalla(1, 10, (region1), captura_inicial)

