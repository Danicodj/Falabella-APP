@echo off
title Actualizador Automatico Sensing + GitHub
cls

echo ==================================================
echo   1. ACTUALIZANDO DATOS DEL SENSING E IHS...
echo ==================================================
python procesar_datos.py

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Hubo un fallo en el procesamiento.
    pause
    exit /b
)

echo.
echo ==================================================
echo   2. SUBIENDO AUTOMATICAMENTE A GITHUB PAGES...
echo ==================================================
git add data_app.json
git commit -m "Auto-update data sensing"
git push origin main

echo.
echo ==================================================
echo   PROCESO TERMINADO Y PUBLICADO EN INTERNET EXITOSAMENTE
echo ==================================================
pause