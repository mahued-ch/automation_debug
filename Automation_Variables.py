import os
import json
from pathlib import Path
import getpass

# Obtener el directorio donde se ejecuta el programa
#ruta = os.getcwd()
ruta_script = os.path.abspath(__file__)  # Ruta completa del script en ejecución
ruta = os.path.dirname(ruta_script)  # Solo el directorio

# Ruta a la subcarpeta "imagenes" dentro del directorio actual
ruta_imagenes = os.path.join(ruta, "imagenes")
ruta_output = os.path.join(ruta, "output")

ruta_directorio = Path(ruta_output)
ruta_directorio.mkdir(parents=True, exist_ok=True)

log_debug = os.path.join(ruta_imagenes, "log_debug.png")
boton_auditoria1 =  os.path.join(ruta_imagenes, "debug_boton_auditoria1.png")
boton_auditoria2  = os.path.join(ruta_imagenes, "debug_boton_auditoria2.png")
boton_flecha_atras  = os.path.join(ruta_imagenes, "boton_flecha_atras.png")
boton_ok_intro  = os.path.join(ruta_imagenes, "boton_ok_intro.png")
boton_cancel  = os.path.join(ruta_imagenes, "boton_cancel.png")
grabar_lista_fichero  = os.path.join(ruta_imagenes, "grabar_lista_fichero.png")
texto_con_tabuladores  = os.path.join(ruta_imagenes, "texto_con_tabuladores.png")
grabar_fichero  = os.path.join(ruta_imagenes, "grabar_fichero.png")
boton_reemplazar  = os.path.join(ruta_imagenes, "boton_reemplazar.png")
boton_permitir  = os.path.join(ruta_imagenes, "boton_permitir.png")
boton_reinicializar  = os.path.join(ruta_imagenes, "boton_reinicializar.png")
boton_seleccion_multiple  = os.path.join(ruta_imagenes, "boton_seleccion_multiple.png")
boton_arriba_abajo  = os.path.join(ruta_imagenes, "boton_arriba_abajo.png")
boton_seleccionar_detalle  = os.path.join(ruta_imagenes, "boton_seleccionar_detalle.png")
boton_leer_auditoria  = os.path.join(ruta_imagenes, "boton_leer_auditoria.png")
texto_critico = os.path.join(ruta_imagenes, "texto_critico.png")
texto_grave = os.path.join(ruta_imagenes, "texto_grave.png")
texto_nocritico_BU4 = os.path.join(ruta_imagenes, "texto_nocritico_BU4.png")
texto_nocritico_AU5 = os.path.join(ruta_imagenes, "texto_nocritico_AU5.png")
texto_nocritico_AUK = os.path.join(ruta_imagenes, "texto_nocritico_AUK.png")
ventana_noresultados = os.path.join(ruta_imagenes, "ventana_noresultados.png")
configuracion = {}
region1 = (0.05, 0.15, 0.40, 0.40) 
region2 = (0.30, 0.30, 0.50, 0.50) 
region3 = (0.30, 0.15, 0.45, 0.75) 
region4 = (0.18, 0.43, 0.57, 0.72) 

running_test = 0 # 1 = si, 0 = no

###--------------------------------------------------------------------------------
def obtener_ruta_config_usuario():
    if os.name == 'nt':  # Windows
        base = os.getenv("APPDATA")
    else:
        base = os.path.expanduser("~/.config")
    carpeta = os.path.join(base, "automation_debug")
    os.makedirs(carpeta, exist_ok=True)
    return os.path.join(carpeta, "config_user.json")

###--------------------------------------------------------------------------------
def carga_config():
    # Ruta del script actual
    ruta_script = os.path.dirname(__file__)

    # Archivos de configuración
    ruta_config_general = os.path.join(ruta_script, "config.json")
    ruta_config_usuario = os.path.join(ruta_script, "config_user.json")

    # Cargar config general
    with open(ruta_config_general, "r", encoding="utf-8") as file:
        config = json.load(file)

    # Asegurar estructura anidada
    config.setdefault("s4p", {})
    config.setdefault("fiori", {})
    config.setdefault("email", {})

    # Cargar config del usuario si existe
    if os.path.exists(ruta_config_usuario):
        with open(ruta_config_usuario, "r", encoding="utf-8") as file:
            config_local = json.load(file)
            config["s4p"].update(config_local.get("s4p", {}))
            config["fiori"].update(config_local.get("fiori", {}))
            config["email"].update(config_local.get("email", {}))
    else:
        # Solicitar datos la primera vez
        print("⚙️  Primera configuración del usuario:")
        config["s4p"]["usuario"] = input("Usuario S4P: ")
        config["s4p"]["password"] = getpass.getpass("Contraseña S4P: ")
        config["fiori"]["usuario"] = input("Usuario Fiori: ")
        config["fiori"]["password"] = getpass.getpass("Contraseña Fiori: ")
        config["email"]["sender"] = input("Correo electrónico Receptor: ")

        # Guardar 
        config_guardar = {
            "s4p": {
                "usuario": config["s4p"]["usuario"],
                "password": config["s4p"]["password"]
            },
            "fiori": {
                "usuario": config["fiori"]["usuario"],
                "password": config["fiori"]["password"]
            },
            "email": {
                "sender": config["email"]["sender"]
            }
        }

        with open(ruta_config_usuario, 'w', encoding='utf-8') as f:
            json.dump(config_guardar, f, indent=4)

    return config

###--------------------------------------------------------------------------------
#def carga_config():
## Cargar archivo de configuración
#    with open(os.path.join(ruta, "config.json"), "r", encoding="utf-8") as file:
#        configura = json.load(file)
#    return configura    

###########################
def limpiar_pantalla():
    if os.name == 'nt':
        os.system('cls')
    else:
        os.system('clear')   

###########################
def borrar_archivos_carpeta_con_prefijo(carpeta, prefijo):
    """
    Borra todos los archivos en una carpeta que comienzan con un prefijo determinado.

    Args:
        carpeta (str): Ruta de la carpeta donde se encuentran los archivos.
        prefijo (str): Prefijo con el que deben comenzar los archivos para ser borrados.
    """
    # Verifica si la carpeta existe
    if os.path.exists(carpeta):
        # Itera sobre todos los archivos en la carpeta
        for archivo in os.listdir(carpeta):
            archivo_path = os.path.join(carpeta, archivo)
            # Verifica si es un archivo y si empieza con el prefijo
            if os.path.isfile(archivo_path) and archivo.lower().startswith(prefijo.lower()):
                os.remove(archivo_path)
                print(f"Borrado: {archivo_path}")
    else:
        print(f"La carpeta {carpeta} no existe.")
        
###########################

def borrar_archivos_carpeta(carpeta):
    p = Path(carpeta)
    if p.exists() and p.is_dir():
        for archivo in p.iterdir():
            if archivo.is_file():
                archivo.unlink()  # Elimina el archivo
                print(f"Borrado: {archivo}")
    else:
        print(f"La carpeta {carpeta} no existe o no es un directorio.")

###########################
