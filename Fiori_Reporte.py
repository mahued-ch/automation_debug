import pyautogui
import time
import os
import cv2
import sys
from datetime import datetime, timedelta
from Automation_Functions import busca_en_pantalla, buscar_todas_ocurrencias, limpiar_pantalla, espera_cambio_pantalla, captura_pantalla
from Automation_Variables import ruta, ruta_output, boton_flecha_atras, boton_ok_intro, grabar_lista_fichero, texto_con_tabuladores, grabar_fichero, boton_reemplazar, boton_permitir, region1

def get_reporte(lcprefijo, lcfecha_amd):
#    Ejecuta el paso de generación de reportes en el sistema Fiori
#    Este va a servir para los críticos y no críticos y para los graves.
  print("Selección de columnas...")
  # Presionar Ctrl-F8
  time.sleep(1)
  pyautogui.keyDown('ctrl')  # Mantiene presionada la tecla Control
  pyautogui.press('f8')       # Presiona F8
  pyautogui.keyUp('ctrl')    # Suelta la tecla Shift

  time.sleep(2)

  # Adiciona todas las columnas
  print("Adicionando columnas...")
  location = busca_en_pantalla(boton_flecha_atras, 1, 10) 
  location = (location[0], location[1] + 15)
  for i in range(8):
      pyautogui.click(location)
      time.sleep(1)

  captura_inicial = captura_pantalla((region1))    
  location = busca_en_pantalla(boton_ok_intro, 1, 10) 
  if location != None:
    pyautogui.click(location)
#    time.sleep(1)
    cambio_detectado = espera_cambio_pantalla(1, 10, (region1), captura_inicial)
    if not cambio_detectado:
      print('fiori_reporte - 1 - No se encontró cambio de pantalla')
      sys.exit()

  captura_inicial = captura_pantalla((region1))    
  print("Presiono Ctrl-Shift-F9 ...")
  # Presionar Ctrl-Shift-F9
  pyautogui.keyDown('ctrl')  
  pyautogui.keyDown('shift') 
  pyautogui.press('f9')      
  pyautogui.keyUp('shift')   
  pyautogui.keyUp('ctrl')  
  print("Esperando Grabar lista fichero...")
  cambio_detectado = espera_cambio_pantalla(1, 10, (region1), captura_inicial)
  if not cambio_detectado:
    print('fiori_reporte - 2 - No se encontró cambio de pantalla')
    sys.exit()

  location = busca_en_pantalla(grabar_lista_fichero, 5, 12) 

  location = pyautogui.locateCenterOnScreen(texto_con_tabuladores, confidence=0.8)
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

  # Presionar Shift-tab
  pyautogui.keyDown('shift') 
  pyautogui.press('tab')      
  pyautogui.keyUp('shift')   

  # Y teclea la ruta del archivo de salida
  pyautogui.typewrite(ruta_output)

  # Esperando botón reemplazar
  print("Esperando Botón Reemplazar...")
  location = busca_en_pantalla(boton_reemplazar, 1, 20) 
  if location != None:
    pyautogui.click(location)

  # Esperando botón reemplazar
  print("Esperando Botón Permitir...")
  location = busca_en_pantalla(boton_permitir, 1, 20) 
  if location != None:
    pyautogui.click(location)

