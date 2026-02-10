@echo off
title Automatizacion de Facturas - FrasCL
echo ==========================================
echo    INICIANDO PROCESO DE FACTURACION
echo ==========================================
echo.

:: Correccion para rutas de red (UNC)
pushd "%~dp0"

echo 1. Verificando Python...
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo [ERROR] Python no esta instalado en este equipo.
    echo Por favor, instala Python 3.12 desde la Microsoft Store o python.org.
    popd
    pause
    exit /b
)

echo 2. Verificando librerias necesarias...
echo (Esto solo tardara unos segundos la primera vez)
python -m pip install pandas openpyxl reportlab --quiet --no-warn-script-location

echo 3. Iniciando el script...
python generar_facturas.py

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Hubo un problema al generar los informes.
    echo Asegurate de que el Excel no este abierto y tenga el nombre correcto.
) else (
    echo.
    echo [EXITO] Informes generados correctamente en la carpeta 'Facturas_Generadas'.
)

popd
echo.
echo Presiona cualquier tecla para cerrar esta ventana...
pause >nul
