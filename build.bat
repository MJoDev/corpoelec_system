@echo off
setlocal

echo ============================================================
echo  SGIP - Build de distribucion portable
echo  Requiere Python 3.9-3.13 (pywebview usa pythonnet)
echo ============================================================

:: pywebview depends on pythonnet which only supports Python <= 3.13.
:: Python 3.14 is not yet supported. We use py -3.13 explicitly.
py -3.13 --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo ERROR: Python 3.13 no encontrado.
    echo Instala Python 3.13 desde https://python.org y vuelve a intentar.
    echo.
    echo Versiones instaladas:
    py --list
    pause & exit /b 1
)

echo Usando: && py -3.13 --version

:: ── Step 1: Install dependencies ────────────────────────────────────────────
echo.
echo [1/3] Instalando dependencias con Python 3.13...
py -3.13 -m pip install pywebview openpyxl reportlab --quiet
if errorlevel 1 (
    echo ERROR al instalar dependencias.
    pause & exit /b 1
)

py -3.13 -m pip install pyinstaller --quiet
if errorlevel 1 (
    echo ERROR al instalar PyInstaller.
    pause & exit /b 1
)

:: ── Step 2: Build ────────────────────────────────────────────────────────────
echo.
echo [2/3] Compilando con PyInstaller (Python 3.13)...
if not exist "dist\SGIP_Portable" mkdir "dist\SGIP_Portable"
py -3.13 -m PyInstaller ^
    --onefile ^
    --windowed ^
    --name "SGIP" ^
    --distpath "dist\SGIP_Portable" ^
    --add-data "sgip.html;." ^
    --add-data "static;static" ^
    --add-data "logo-servicio-corpoelec.im1761941416671im.avif;." ^
    main.py

if errorlevel 1 (
    echo.
    echo ERROR: La compilacion fallo. Revise los mensajes anteriores.
    pause & exit /b 1
)

echo.
echo [3/3] Verificando distribucion...

echo.
echo ============================================================
echo  SGIP compilado exitosamente.
echo  Ejecutable: dist\SGIP_Portable\SGIP.exe
echo.
echo  Copie la carpeta dist\SGIP_Portable\ al pendrive.
echo  Al primer arranque se crean automaticamente:
echo    sgip_data.db     (base de datos SQLite)
echo    assets\images\   (imagenes de materiales)
echo.
echo ============================================================
echo.
pause
