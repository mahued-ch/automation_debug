@echo off
setlocal

echo ======================================
echo   Ejecutando Automation_Debug_Fiori.py
echo ======================================
call venv\Scripts\activate
python Automation_Debug_Fiori.py

echo ======================================
echo   Finalizado. Presiona una tecla para salir.
echo ======================================
pause
endlocal
