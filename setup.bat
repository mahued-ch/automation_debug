@echo off
setlocal

echo ======================================
echo   Creando entorno virtual (venv)...
echo ======================================
python -m venv venv

if exist venv\Scripts\activate (
    echo ======================================
    echo   Activando entorno virtual...
    echo ======================================
    call venv\Scripts\activate

    echo ======================================
    echo   Instalando dependencias...
    echo ======================================
    pip install --upgrade pip
    pip install -r requirements.txt

    echo ======================================
    echo   Entorno configurado correctamente. ✅
    echo ======================================
) else (
    echo ❌ Error: No se pudo crear el entorno virtual.
)

pause
endlocal
