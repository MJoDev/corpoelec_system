@echo off
:: Run SGIP in development mode (no compilation needed).
:: Uses Python 3.13 — required by pywebview/pythonnet.

py -3.13 --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python 3.13 no encontrado.
    echo Instala Python 3.13 desde https://python.org
    pause & exit /b 1
)

:: Install deps if not present
py -3.13 -m pip show pywebview >nul 2>&1
if errorlevel 1 (
    echo Instalando dependencias...
    py -3.13 -m pip install pywebview openpyxl reportlab --quiet
)

echo Iniciando SGIP...
py -3.13 main.py
