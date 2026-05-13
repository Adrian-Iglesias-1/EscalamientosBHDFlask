@echo off
chcp 65001 >nul
title EscalamientosApp - Deteniendo...

echo ============================================
echo    EscalamientosApp - Deteniendo...
echo ============================================

echo [1/2] Cerrando servidor Flask...
if exist "backend\server.pid" (
    set /p PID=<backend\server.pid
    taskkill /F /PID %PID% 2>nul
    del backend\server.pid
)

echo [2/2] Cerrando todos los procesos Python...
taskkill /F /IM python.exe 2>nul
taskkill /F /IM pythonw.exe 2>nul

echo.
echo Servidores detenidos correctamente.
timeout /t 2 /nobreak >nul
exit /b
