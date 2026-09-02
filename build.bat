@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

REM ============================================================
REM  Compilacion de CruceLecturas
REM  Dejar este archivo en la MISMA carpeta que cruce_lecturas.py
REM  Uso: doble clic, o "build.bat deploy" para copiar a la red
REM ============================================================

set "SCRIPT=cruce_lecturas.py"
set "NOMBRE=CruceLecturas"
set "ICONO=icono.ico"
set "RED=\\192.168.10.4\s\Gas Natural Lectura 2019-2020\CruceLecturas_verificadores"

echo.
echo ============================================
echo   Compilando %NOMBRE%
echo ============================================
echo.

REM ---- 1. Verificaciones previas -----------------------------
if not exist "%SCRIPT%" (
    echo [ERROR] No se encuentra %SCRIPT% en esta carpeta.
    echo         Deja build.bat junto al archivo .py
    goto :fin
)

python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python no esta instalado o no esta en el PATH.
    goto :fin
)

REM Leer la version desde el .py, para dejarla registrada
for /f "tokens=2 delims==" %%v in ('findstr /b /c:"VERSION = " "%SCRIPT%"') do (
    set "VER=%%v"
)
set "VER=!VER:"=!"
set "VER=!VER: =!"
echo   Version detectada: !VER!
echo.

REM ---- 2. Dependencias ---------------------------------------
echo [1/4] Verificando dependencias...
python -c "import openpyxl" 2>nul || (
    echo       Instalando openpyxl...
    python -m pip install --quiet openpyxl || goto :error_dep
)
python -c "import PyInstaller" 2>nul || (
    echo       Instalando pyinstaller...
    python -m pip install --quiet pyinstaller || goto :error_dep
)
echo       OK
echo.

REM ---- 3. Limpieza -------------------------------------------
echo [2/4] Limpiando compilacion anterior...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
if exist "%NOMBRE%.spec" del /q "%NOMBRE%.spec"
echo       OK
echo.

REM ---- 4. Compilacion ----------------------------------------
echo [3/4] Compilando (puede tardar un minuto)...
echo.

set "OPT_ICONO="
if exist "%ICONO%" (
    set "OPT_ICONO=--icon=%ICONO% --add-data %ICONO%;."
) else (
    echo       [AVISO] No se encontro %ICONO%. Se compila sin icono.
)

pyinstaller --windowed --noconfirm ^
    --name %NOMBRE% ^
    --contents-directory data ^
    !OPT_ICONO! ^
    "%SCRIPT%"

if errorlevel 1 goto :error_build
if not exist "dist\%NOMBRE%\%NOMBRE%.exe" goto :error_build
echo.

REM ---- 5. Resultado ------------------------------------------
echo [4/4] Listo.
echo.
echo   Ejecutable: %CD%\dist\%NOMBRE%\%NOMBRE%.exe
echo   Version:    !VER!
echo.

REM ---- 6. Despliegue opcional --------------------------------
if /i "%~1"=="deploy" goto :deploy
echo   Para copiar a la red, ejecuta:  build.bat deploy
echo.
goto :fin

:deploy
echo ============================================
echo   Despliegue a la red
echo ============================================
echo   Destino: %RED%\App
echo.
if not exist "%RED%" (
    echo [ERROR] No se puede acceder a %RED%
    echo         Verifica la conexion de red y los permisos.
    goto :fin
)
echo   IMPORTANTE: nadie debe tener la app abierta.
echo.
set /p "OK=Continuar? (S/N): "
if /i not "!OK!"=="S" goto :fin

REM Respaldo de la version anterior, para poder volver atras
if exist "%RED%\App" (
    if exist "%RED%\App_ANTERIOR" rmdir /s /q "%RED%\App_ANTERIOR"
    move "%RED%\App" "%RED%\App_ANTERIOR" >nul
    if errorlevel 1 (
        echo [ERROR] No se pudo mover la version anterior.
        echo         Probablemente alguien tiene la app abierta.
        goto :fin
    )
    echo   Version anterior guardada en App_ANTERIOR
)

xcopy "dist\%NOMBRE%" "%RED%\App\" /e /i /q /y >nul
if errorlevel 1 (
    echo [ERROR] Fallo la copia. Restaurando version anterior...
    if exist "%RED%\App" rmdir /s /q "%RED%\App"
    if exist "%RED%\App_ANTERIOR" move "%RED%\App_ANTERIOR" "%RED%\App" >nul
    goto :fin
)

echo.
echo   Desplegado correctamente en %RED%\App
echo   La config.json no se toco.
echo.
echo   Probalo desde otra PC antes de borrar App_ANTERIOR.
echo.
goto :fin

:error_dep
echo.
echo [ERROR] No se pudieron instalar las dependencias.
echo         Si hay proxy corporativo, instalalas a mano.
goto :fin

:error_build
echo.
echo [ERROR] Fallo la compilacion.
echo         Para ver el detalle, compila sin --windowed:
echo         pyinstaller --name %NOMBRE%_debug "%SCRIPT%"
goto :fin

:fin
echo.
pause
endlocal
