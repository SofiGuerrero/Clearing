@echo off
echo Instalando dependencias...
pip install flask pywin32 pandas openpyxl --quiet
echo.
echo Iniciando SAP Clearing Dashboard...
python "%~dp0server_standalone.py"
pause
