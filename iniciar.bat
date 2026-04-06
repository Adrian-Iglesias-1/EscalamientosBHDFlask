@echo off
chcp 65001 >nul
title EscalamientosApp - Iniciando

echo ============================================
echo EscalamientosApp - Iniciando
echo ============================================
echo.

:: Ir a la carpeta del script (donde esté ubicado)
cd /d "%~dp0"
if not exist "backend\app.py" (
    echo ERROR: No se encontro backend\app.py
    echo.
    echo Asegurate de descomprimir el ZIP completamente.
    echo El archivo iniciar.bat debe estar en la carpeta raiz del proyecto.
    pause
    exit /b 1
)

:: Crear acceso directo en el escritorio (solo si no existe)
if not exist "%USERPROFILE%\Desktop\EscalamientosApp.lnk" (
    echo [0/4] Creando acceso directo en el escritorio...
    
    >"%TEMP%\CreateShortcut.vbs" (
        echo Set WshShell = WScript.CreateObject^("WScript.Shell"^)
        echo strDesktop = WshShell.SpecialFolders^("Desktop"^)
        echo Set oShortcut = WshShell.CreateShortcut^(strDesktop ^& "\EscalamientosApp.lnk"^)
        echo oShortcut.TargetPath = "%~f0"
        echo oShortcut.WorkingDirectory = "%~dp0"
        echo oShortcut.IconLocation = "%SystemRoot%\System32\shell32.dll,15"
        echo oShortcut.Save
    )
    cscript //nologo "%TEMP%\CreateShortcut.vbs" >nul 2>&1
    del "%TEMP%\CreateShortcut.vbs" >nul 2>&1
    
    if exist "%USERPROFILE%\Desktop\EscalamientosApp.lnk" (
        echo     Acceso directo creado!
    ) else (
        echo     No se pudo crear acceso directo.
    )
    echo.
)

:: Detectar Python
echo [1/4] Detectando Python...

py --version >nul 2>&1
if not errorlevel 1 (
    set "PYTHON_CMD=py"
    echo     Encontrado: py
    goto python_ok
)

python --version >nul 2>&1
if not errorlevel 1 (
    set "PYTHON_CMD=python"
    echo     Encontrado: python
    goto python_ok
)

python3 --version >nul 2>&1
if not errorlevel 1 (
    set "PYTHON_CMD=python3"
    echo     Encontrado: python3
    goto python_ok
)

echo.
echo ERROR: No se encontro Python instalado.
echo.
echo Instala Python desde el portal de la empresa
echo Asegurate de marcar "Add Python to PATH"
pause
exit /b 1

:python_ok
echo.

:: Entrar a backend y crear venv limpio
cd backend
echo [2/4] Verificando entorno virtual...

:: SIEMPRE eliminar venv existente para evitar problemas de rutas
if exist "venv" (
    echo     Eliminando entorno virtual anterior...
    rd /s /q "venv" >nul 2>&1
)

echo     Creando entorno virtual...
%PYTHON_CMD% -m venv venv
if errorlevel 1 (
    echo.
    echo ERROR: No se pudo crear el entorno virtual.
    pause
    exit /b 1
)
echo     Entorno virtual creado.
echo.

:: Instalar dependencias
echo [3/4] Instalando dependencias...
echo     (La primera vez puede tardar unos minutos)

call venv\Scripts\pip install --upgrade pip --quiet 2>nul
call venv\Scripts\pip install -r requirements.txt --quiet 2>nul
if errorlevel 1 (
    call venv\Scripts\pip install -r requirements.txt
    if errorlevel 1 (
        echo.
        echo ERROR: No se pudieron instalar las dependencias.
        pause
        exit /b 1
    )
)
echo     Dependencias listas.
echo.

:: Verificar si servidor ya corre
echo [4/4] Iniciando servidor...
netstat -an 2>nul | findstr ":5000" | findstr "LISTENING" >nul 2>&1
if not errorlevel 1 (
    echo     Servidor YA estaba corriendo.
    start http://localhost:5000
    echo.
    echo Esta ventana se cerrara en 3 segundos...
    timeout /t 3 /nobreak >nul
    exit
)

:: Probar si el servidor inicia
echo     Probando servidor...
call venv\Scripts\python.exe -c "from app import app; print('OK')" >nul 2>&1
if errorlevel 1 (
    echo     ERROR: No se pudo iniciar el servidor.
    echo.
    echo Revisando dependencias...
    call venv\Scripts\pip list
    pause
    exit /b 1
)

:: Iniciar en segundo plano
echo     Iniciando en segundo plano...
start "EscalamientosApp Server" venv\Scripts\pythonw.exe app.py
timeout /t 2 /nobreak >nul
start http://localhost:5000
echo.
echo ----------------------------------------
echo     Servidor iniciado correctamente!
echo ----------------------------------------
echo.
echo La aplicacion se abrio en el navegador.
echo Esta ventana se cerrara en 3 segundos...
timeout /t 3 /nobreak >nul
exit
