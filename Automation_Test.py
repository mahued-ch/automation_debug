from Automation_Functions import convert_to_excel, convert_to_excel_fiori, abrir_correo_outlook
from Automation_Variables import ruta_output
import os

# Ejemplo de uso:
if __name__ == "__main__":
#    resultado = convert_to_excel(r"C:\OD\OneDrive - Tiendas Chedraui S.A. de C.V\Documents\Python\.vscode\Automation_Debug\output\analisis\s4p_20250202.xls")
#    resultado = convert_to_excel(r"C:\Users\m_ahu\OneDrive - Tiendas Chedraui S.A. de C.V\Documents\Python\.vscode\Automation_Debug\output\analisis\s4p_20250202.xls")
    resultado = convert_to_excel_fiori(r"C:\OD\OneDrive - Tiendas Chedraui S.A. de C.V\Documents\Python\.vscode\Automation_Debug\output\analisis\Fiori_criticos_20250202.xls")

    print('Archivo_Salida: ', resultado["archivo_salida"])  # Nombre del archivo generado
    print('Analisis: ', resultado["analisis"])  # Datos analizados en formato diccionario

    archivos = [resultado["archivo_salida"]]  # Lista de archivos a adjuntar    
    
    prefijo = os.path.splitext(os.path.basename(resultado["archivo_salida"]))[0]    
    ruta = os.path.dirname(resultado["archivo_salida"])
    
    abrir_correo_outlook(resultado, "21.10.2024", "27.10.2024", "mahued@chedraui.com.mx", archivos, ruta, prefijo, ['.png', '.docx'])    
   