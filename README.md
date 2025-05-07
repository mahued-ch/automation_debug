
# Proyecto de Automatización

Este proyecto contiene dos programas principales para la automatización de tareas en diferentes sistemas.

## Estructura del Proyecto

```plaintext
Automation_Debug/
├── setup.bat                       # Script para crear el entorno virtual e instalar dependencias
├── run_fiori.bat                   # Script para ejecutar Automation_Debug_Fiori.py
├── run_sap.bat                     # Script para ejecutar Automation_Debug_SAP.py
├── Automation_Debug_Fiori.py       # Primer programa de automatización
├── Automation_Debug_SAP.py         # Segundo programa de automatización
├── Automation_Convert_2_Excel.py   # Script para convertir la salida a excel
├── Automation_Functions            # Script de funciones
├── Automation_Genera_Docto         # Script para generar el .docx
├── Automation_Variables            # Declaración de variables
├── requirements.txt                # Paquetes necesarios para el proyecto
└── .gitignore                      # Archivos y carpetas ignorados por Git
```

## Requisitos

Antes de ejecutar cualquiera de los programas, debes asegurarte de tener **Python 3** instalado. Si aún no tienes Python, puedes descargarlo desde [aquí](https://www.python.org/downloads/).

### Instalación

1. **Clona o descarga el proyecto** en tu máquina.

2. **Ejecuta el archivo `setup.bat`** para crear un entorno virtual e instalar las dependencias:
   
   Solo haz doble clic en el archivo `setup.bat`. Esto:
   - Crea un entorno virtual en la carpeta `venv/`
   - Instala las dependencias necesarias desde `requirements.txt`

3. **Elige qué programa ejecutar:**
   
   Después de haber configurado el entorno virtual, ejecuta uno de los siguientes scripts:

   - **Para ejecutar el programa Fiori:**
     - Haz doble clic en `run_fiori.bat`
   
   - **Para ejecutar el programa S4P:**
     - Haz doble clic en `run_s4p.bat`

   Esto ejecutará el script correspondiente con las dependencias necesarias activadas.

## Estructura de los Programas

- **`Automation_Debug_Fiori.py`**: Este script realiza tareas de automatización relacionadas con Fiori.
  
- **`Automation_Debug_SAP.py`**: Este script realiza tareas de automatización relacionadas con S4P.

## Notas

- Los programas utilizan un entorno virtual para gestionar las dependencias, por lo que siempre es recomendable usar el archivo `setup.bat` para asegurarse de que todo esté bien configurado.
- Si tienes problemas para ejecutar alguno de los scripts, asegúrate de que todos los paquetes estén correctamente instalados ejecutando:

   ```bash
   pip install -r requirements.txt
   ```

## Contribuciones

Si encuentras algún detalle en la ejecución o quieres hacer una sugerecia manda tus comentarios a mi correo.
