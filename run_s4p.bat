@echo off
setlocal

echo ======================================
echo   Ejecutando Automation_Debug_SAP.py
echo ======================================
call venv\Scripts\activate
python Automation_Debug_SAP.py

echo ======================================
echo   Finalizado. Presiona una tecla para salir.
echo ======================================
pause
endlocal
