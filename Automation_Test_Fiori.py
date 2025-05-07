import pyautogui
import time
import os
import cv2
import sys
from datetime import datetime, timedelta
from pywinauto import Application
from Fiori_Reporte import get_reporte
from Automation_Variables import ruta, ruta_output, boton_flecha_atras, boton_ok_intro, grabar_lista_fichero, texto_con_tabuladores, grabar_fichero, boton_reemplazar, boton_permitir 
from Automation_Variables import boton_cancel, boton_auditoria1, boton_auditoria2, boton_reinicializar, log_debug, boton_seleccion_multiple, boton_arriba_abajo, boton_seleccionar_detalle
from Automation_Variables import boton_leer_auditoria, texto_grave, texto_nocritico_AU5, texto_nocritico_AUK, texto_nocritico_BU4, texto_critico, texto_con_tabuladores, ventana_noresultados
from Automation_Genera_Docto import Genera_Documento
from Automation_Variables import borrar_archivos_carpeta_con_prefijo, limpiar_pantalla, running_test 
from Automation_Variables import configuracion, carga_config
from Automation_Functions import convert_to_excel_fiori, abrir_correo_outlook
from Automation_SAP import start_sap_logon, login_to_sap, close_window, close_sap_logon

###--------------------------------------------------------------------------------
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

from pywinauto import Application

###--------------------------------------------------------------------------------

def listar_elementos_de_ventanas(app):
    """
    Lista todos los elementos dentro de cada ventana detectada en la aplicación.
    
    :param app: Aplicación conectada con pywinauto.
    """
    try:
        ventanas = app.windows()
        for ventana in ventanas:
            print(f"\n🔹 Ventana: {ventana.window_text()} \n{'-'*50}")
            try:
            # Obtener todos los elementos de la ventana
                elementos = ventana.descendants()
            
                for elemento in elementos:
                    print(f"🖥️ {elemento.window_text()} - {elemento.control_type()}")
        
            except Exception as e:
                print(f"⚠️ No se pudo obtener los elementos de la ventana: {e}")
    except Exception as e:
         print(f"⚠️ No se pudo obtener los elementos de la ventana: {e}")

###--------------------------------------------------------------------------------

def listar_elementos_treelist(app):
    """
    Lista todos los elementos dentro del 'SAP's Advanced Treelist'.
    
    :param app: Aplicación conectada con pywinauto.
    """
    # Obtener todas las ventanas de la aplicación
    ventanas = app.windows()

    for ventana in ventanas:
        try:
            # Buscar el Treelist dentro de los elementos descendientes de la ventana
            for elemento in ventana.descendants():
                
                auto_id = getattr(elemento.element_info, "automation_id", "N/A")  # Si no tiene auto_id, pone "N/A"
                print(f"Texto: {elemento.window_text()}, ID: {auto_id}, Tipo: {elemento.element_info.control_type}")
                
                if elemento.window_text() == "Control de encabezado" and elemento.element_info.control_type == "Pane":
                    
                    print(f"\n🔹 Elementos dentro del Treelist en la ventana '{ventana.window_text()}':\n{'-'*60}")

                    # Obtener y listar todos los elementos hijos dentro del Treelist
                    elementos = elemento.descendants()
                    
                    # 👉 DEPURACIÓN: Mostrar cuántos elementos se encontraron
                    print(f"📌 Número de elementos encontrados en el Treelist: {len(elementos)}")
                    
                    if not elementos:
                        print("⚠️ El Treelist no tiene elementos detectables.")
                        return
                    
                    for item in elementos:
                        print(f"🖥️ {item.window_text()} - {item.element_info.control_type}")

                    return  # Salimos después de encontrar el Treelist

        except Exception as e:
            print(f"⚠️ Error al acceder al Treelist: {e}")

###--------------------------------------------------------------------------------

limpiar_pantalla()

# Borramos todos los archivos de la ruta de salida
if running_test == 0:
    
#  for i in range(15):
#    print(f"Mostramos la página {i}.")

#    hubo_click = False
#    time.sleep(2)  # Pequeña pausa entre clics

# Pruebas
# Conectar a la ventana de la aplicación (ajustar el nombre según tu sistema)
#    windows = app.windows()

#    app = Application(backend="uia").connect(title_re="Evalu*", timeout=10)
#    sap_window = app.window(title_re="Evalu*")
#    print(sap_window.print_control_identifiers())
#    limpiar_pantalla()
#    listar_elementos_de_ventanas(app)

# Conectar a la aplicación
    app = Application(backend="uia").connect(title_re="Evaluación del log de auditoría de seguridad", timeout=10)

# Llamar a la función para listar los elementos del Treelist
    listar_elementos_treelist(app)

# Obtener el panel "SAP's Advanced Treelist"
#    tree_list = sap_window.child_window(title="SAP's Advanced Treelist", auto_id="100", control_type="Pane")
#    items = tree_list.descendants()

# Mostrar los nombres de los ítems encontrados
#    for item in items:
#        print(f"Elemento encontrado: {item.window_text()} - {item.control_type()}")    

#    listar_elementos_de_ventanas(sap_window)
    
#    windows = app.windows()
 # Obtener la ventana principal
#    for win in windows:
#        print(f"Ventana encontrada: {win.window_text()}")
        
# Buscar el checkbox de "BU 5" por su texto
#    checkbox = window.child_window(title="BU 5", control_type="CheckBox")

# Hacer clic en el checkbox
#    checkbox.click()

 #   print("✅ Click en BU 5 exitoso.")
    
#

########################################################################
# Fin del programa
########################################################################

print("Fin del Programa")
sys.exit()  
