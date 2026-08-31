@echo off
title Gestor Automatizado de Reclamaciones B2B - Home Impex
echo ========================================================
echo   Iniciando Gestor Automatizado de Reclamaciones B2B...
echo ========================================================
echo.

cd /d "%~dp0"

REM Verificar si Python esta instalado
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] No se encontro Python en el sistema.
    echo Por favor, instala Python desde python.org y asegurate de marcar "Add Python to PATH".
    pause
    exit /b
)

echo.
echo ========================================================
echo   ACCESO LOCAL Y EN RED DE OFICINA:
echo   - Tu propio PC:       http://localhost:8501
echo   - Companeros en red:  http://192.168.1.35:8501
echo ========================================================
echo.

echo Abriendo aplicacion en tu navegador...
start "" http://localhost:8501

REM Ejecutar Streamlit permitiendo conexiones de la red local
python -m streamlit run app.py --server.address 0.0.0.0 --server.port 8501 --server.headless false

pause
