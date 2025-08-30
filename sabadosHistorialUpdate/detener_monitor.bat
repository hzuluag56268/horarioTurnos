@echo off
echo ========================================
echo    DETENIENDO MONITOR DE TROP
echo ========================================
echo.
echo Buscando procesos de Python...
echo.

REM Buscar y terminar procesos de Python que ejecuten trop_monitor.py
tasklist /FI "IMAGENAME eq python.exe" /FO CSV | findstr /I "python.exe" > nul
if %errorlevel% equ 0 (
    echo Procesos de Python encontrados:
    tasklist /FI "IMAGENAME eq python.exe" /FO TABLE
    echo.
    echo Terminando procesos de Python...
    taskkill /F /IM python.exe
    echo.
    echo Procesos terminados exitosamente.
) else (
    echo No se encontraron procesos de Python ejecutándose.
)

echo.
echo Presiona cualquier tecla para cerrar...
pause > nul 