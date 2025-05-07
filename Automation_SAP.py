import os
import sys
import time
import psutil
import subprocess
from pywinauto import Application
from Automation_Functions import limpiar_pantalla, espera_cambio_pantalla, captura_pantalla, copiar_imagen_al_clipboard
from Automation_Variables import configuracion, carga_config, region1, region2
from pywinauto.keyboard import send_keys

# Ruta de SAP Logon, ajusta si es diferente
#SAP_LOGON_PATH = config["sap"]["gui_path"] #r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"

# Configuración del sistema y credenciales
#SAP_SYSTEM_NAME = "NOMBRE_DEL_SISTEMA"  # Cambia por el sistema SAP al que deseas conectarte
#SAP_USER = "usuario"
#SAP_PASSWORD = "contraseña"

def is_sap_running():
    """Verifica si SAP Logon ya está en ejecución."""
    for proc in psutil.process_iter(attrs=["name", "pid"]):
        if "saplogon.exe" in proc.info["name"].lower():
            return proc.info["pid"]  # Devuelve el PID del proceso
    return None

def start_sap_logon(configuracion):
    """Inicia SAP Logon si no está abierto."""
    if not is_sap_running():
        captura_inicial = captura_pantalla(region2) #(1000, 540, 1500, 700))    
        
        subprocess.Popen(configuracion["sap"]["gui_path"], shell=True)
#        time.sleep(5)  # Esperar a que abra SAP Logon
#        cambio_detectado = espera_cambio_pantalla(2, 10, (1000, 540, 1500, 700), captura_inicial)
        cambio_detectado = espera_cambio_pantalla(2, 10, (region2), captura_inicial)
        if not cambio_detectado:
          print('start_sap_logon - No se encontró ventana')
          sys.exit()

def close_window():
    """Cierra la ventana"""
    send_keys("% {F4}")  # Simula ALT + F4
    send_keys("{TAB}")  # TAB
    send_keys("{ENTER}")  # Enter

def close_sap_logon():
    """Cierra SAP Logon si está abierto."""
    pid = is_sap_running()
    if pid:
        process = psutil.Process(pid)
        process.terminate()  # Cierra SAP Logon
        print("SAP Logon cerrado.")
    else:
        print("SAP Logon no estaba abierto.")
        
def login_to_sap(sistema, usuario, password):
    """Automatiza el login a SAP Logon."""
    
#    time.sleep(3)
    app = Application(backend="uia").connect(title_re="SAP Logon.*", timeout=10)
    sap_logon = app.window(title_re="SAP Logon.*")

    captura_inicial = captura_pantalla((region1))    
    sistema_item = sap_logon.child_window(title=sistema, control_type="ListItem").wrapper_object()
    sistema_item.click_input(double=True)
#    time.sleep(3)
    cambio_detectado = espera_cambio_pantalla(2, 10, (region1), captura_inicial)
    if not cambio_detectado:
      print('login_to_sap - 1 - No se encontró ventana de SAP')
      sys.exit()
    
    # Ingresar usuario y contraseña
#    time.sleep(2)
    send_keys(usuario, pause=0.1)
    send_keys("{TAB}", pause=0.1)  # Simula la tecla TAB    
    send_keys(password, pause=0.1)    
    captura_inicial = captura_pantalla((region1))    
    send_keys("{ENTER}", pause=0.1)  # Simula la tecla ENTER    
    cambio_detectado = espera_cambio_pantalla(2, 10, (region1), captura_inicial)
    if not cambio_detectado:
      print('login_to_sap - 2 - No se encontró ventana de SAP')
      sys.exit()
#    time.sleep(5)
    send_keys("% {SPACE} x")  # Simula ALT + ESPACIO + X
    time.sleep(1)
    send_keys("sm20n", pause=0.1)    
    captura_inicial = captura_pantalla((region1))    
    send_keys("{ENTER}", pause=0.1)  # Simula la tecla ENTER    
#    time.sleep(2)
    cambio_detectado = espera_cambio_pantalla(2, 10, (region1), captura_inicial)
    if not cambio_detectado:
      print('login_to_sap - 3 - No se encontró ventana de SAP')
      sys.exit()

def main():
    limpiar_pantalla()
    configuracion = carga_config()
    start_sap_logon(configuracion)
    login_to_sap(configuracion["sap"]["system_s4p"], configuracion["s4p"]["usuario"], configuracion["s4p"]["password"])
    close_window()
    close_sap_logon()

if __name__ == "__main__":
    main()
