@echo off
setlocal
cd /d "%~dp0"

echo ============================================================
echo  SGIP - Build portable compatible con Windows XP 32 bits
echo  Python 3.4.4 x86  +  Flask 1.x  +  PyInstaller 3.3.1
echo ============================================================
echo.

set PY_TAG=-3.4-32

py %PY_TAG% --version >nul 2>&1
if errorlevel 1 goto :no_python34

echo Usando Python: && py %PY_TAG% --version
echo Arquitectura: 32 bits - Windows XP compatible
echo.

rem ---- Step 1: dependencias ----
echo [1/4] Instalando dependencias compatibles con Python 3.4...

rem  Flask<2 resuelve automaticamente jinja2/werkzeug compatibles con Python 3.4
py %PY_TAG% -m pip install "flask<2.0" "markupsafe<2.0" --quiet
if errorlevel 1 goto :err_dep

rem  openpyxl 2.6.x requiere Python 3.5+, el maximo para 3.4 es 2.5.x
py %PY_TAG% -m pip install "openpyxl<2.6" --quiet
if errorlevel 1 goto :err_dep

rem  reportlab requiere MSVC 2010 para compilar; si falla el PDF no estara disponible
py %PY_TAG% -m pip install "reportlab<3.6" --only-binary :all: --quiet
if errorlevel 1 (
    echo [AVISO] reportlab no se pudo instalar - exportacion PDF no disponible.
    echo         Instala Visual C++ 2010 Redistributable si necesitas PDF.
)

rem  pyinstaller==3.3.1 tira de pypiwin32 que requiere pywin32>=223 (no existe para Python 3.4)
rem  Soluccion: instalar sin deps y usar pywin32-ctypes (Python puro, sin compilar)
py %PY_TAG% -m pip install "pyinstaller==3.3.1" --no-deps --quiet
if errorlevel 1 goto :err_dep

py %PY_TAG% -m pip install "pefile" "altgraph" "pywin32-ctypes" --quiet
if errorlevel 1 goto :err_dep

echo Dependencias instaladas correctamente.

rem ---- Step 3: compilar ----
echo.
echo [3/4] Compilando con PyInstaller 3.3.1 ...

if not exist "dist\SGIP_XP" mkdir "dist\SGIP_XP"

rem  Cerrar proceso anterior si esta corriendo y limpiar exe antiguo
taskkill /f /im "SGIP.exe" >nul 2>&1
if exist "dist\SGIP_XP\SGIP.exe" (
    attrib -r "dist\SGIP_XP\SGIP.exe" >nul 2>&1
    del /f /q "dist\SGIP_XP\SGIP.exe"
)

rem  Verificar python34.dll antes de compilar (el spec lo resuelve automaticamente)
set PYTHON34_DLL=C:\Windows\SysWOW64\python34.dll
if not exist "%PYTHON34_DLL%" set PYTHON34_DLL=C:\Windows\System32\python34.dll
if not exist "%PYTHON34_DLL%" (
    echo ERROR: python34.dll no encontrado en System32 ni SysWOW64.
    pause & exit /b 1
)

rem  Usar SGIP.spec para controlar el build (incluye filtro de mozglue.dll)
py %PY_TAG% -m PyInstaller SGIP.spec --distpath "dist\SGIP_XP"

if errorlevel 1 goto :err_build

rem ---- Step 4: listo ----
echo.
echo [4/4] Distribucion lista.
echo.
echo ============================================================
echo  Compilado exitosamente - 32 bits / Windows XP compatible
echo.
echo  Ejecutable: dist\SGIP_XP\SGIP.exe
echo.
echo  En XP: se abre una ventana de control y el navegador
echo         apuntando a http://127.0.0.1:5128
echo.
echo  Para crear el instalador ejecuta: instalador.bat
echo ============================================================
echo.
pause
exit /b 0

:no_python34
echo.
echo ERROR: Python 3.4.4 de 32 bits no encontrado.
echo.
echo Descarga el instalador desde:
echo   https://www.python.org/ftp/python/3.4.4/python-3.4.4.msi
echo Instala y vuelve a ejecutar este script.
echo.
echo Versiones instaladas actualmente:
py --list
echo.
pause
exit /b 1

:err_dep
echo.
echo ERROR: Fallo al instalar una dependencia.
echo Verifica tu conexion a internet y vuelve a intentar.
echo.
pause
exit /b 1

:err_build
echo.
echo ERROR: La compilacion fallo. Revisa los mensajes anteriores.
echo.
pause
exit /b 1
