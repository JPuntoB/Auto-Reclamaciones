@echo off
echo ==============================================
echo Iniciando el Gestor de Reclamaciones B2B...
echo ==============================================
echo.
echo Instalando dependencias necesarias (esto puede tardar unos segundos)...
pip install -r requirements.txt
echo.
echo Iniciando la aplicacion web...
streamlit run app.py
pause
