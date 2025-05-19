@echo off
REM === Script para actualizar el proyecto desde GitHub ===

REM Cambiar al directorio del proyecto (ajusta la ruta si es necesario)
cd /d "%~dp0"

echo.
echo ========================================
echo  Actualizando el repositorio desde GitHub
echo ========================================
echo.

git pull

IF %ERRORLEVEL% NEQ 0 (
    echo.
    echo Error al actualizar. Verifica que Git esté instalado y configurado.
) ELSE (
    echo.
    echo Proyecto actualizado correctamente.
)

pause
