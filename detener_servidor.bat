@echo off
chcp 65001 >nul
title EscalamientosApp - Deteniendo...

echo ============================================
echo    EscalamientosApp - Deteniendo...
echo ============================================
echo.

echo [1/2] Buscando proceso en puerto 5000...
set "APP_PID="
for /f "tokens=5" %%a in ('netstat -ano 2^>nul ^| findstr ":5000 " ^| findstr "LISTENING"') do (
    if not defined APP_PID set "APP_PID=%%a"
)

echo [2/2] Cerrando servidor...
if defined APP_PID (
    taskkill /F /PID %APP_PID% 2>nul
    echo     Proceso PID %APP_PID% detenido.
) else (
    echo     No se encontro servidor activo en puerto 5000.
)

:: Limpiar PID file si existe
if exist "backend\server.pid" del "backend\server.pid" 2>nul

echo.
echo Servidor detenido correctamente.
timeout /t 2 /nobreak >nul
exit /b
