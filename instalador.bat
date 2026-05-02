@echo off
setlocal
cd /d "%~dp0"

echo ============================================================
echo  SGIP - Crear instalador con Inno Setup
echo ============================================================

:: Verificar que el exe ya fue compilado
if not exist "dist\SGIP_XP\SGIP.exe" (
    echo ERROR: dist\SGIP_XP\SGIP.exe no existe.
    echo Ejecuta primero build.bat para compilar.
    pause & exit /b 1
)

:: Buscar Inno Setup en rutas comunes
set ISCC=""
if exist "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" set ISCC="C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
if exist "C:\Program Files\Inno Setup 6\ISCC.exe"       set ISCC="C:\Program Files\Inno Setup 6\ISCC.exe"
if exist "C:\Program Files (x86)\Inno Setup 5\ISCC.exe" set ISCC="C:\Program Files (x86)\Inno Setup 5\ISCC.exe"
if exist "C:\Program Files\Inno Setup 5\ISCC.exe"       set ISCC="C:\Program Files\Inno Setup 5\ISCC.exe"

if %ISCC%=="" (
    echo ERROR: Inno Setup no encontrado.
    echo Descargalo en: https://jrsoftware.org/isdl.php
    pause & exit /b 1
)

if not exist "dist\installer" mkdir "dist\installer"

echo Compilando instalador...
%ISCC% sgip_setup.iss

if errorlevel 1 (
    echo ERROR al crear el instalador.
    pause & exit /b 1
)

echo.
echo ============================================================
echo  Instalador creado: dist\installer\SGIP_Setup_v1.0.0_XP32.exe
echo ============================================================
echo.
pause
