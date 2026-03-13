@echo off
REM Scheduler automatico para Reportes Telcel - Windows Task Scheduler
REM Este archivo debe ejecutarse como Administrator

echo.
echo ========================================
echo Configurador de Automatizacion - Telcel
echo ========================================
echo.

REM Variables
setlocal enabledelayedexpansion
set SCRIPT_PATH=%~dp0generar_reportes.py
set SCRIPT_DIR=%~dp0
set TASK_NAME=TelcelReportesAuto_9AM
set HORA=09:00:00

echo [*] Ruta del script: !SCRIPT_PATH!
echo [*] Directorio: !SCRIPT_DIR!
echo.

REM Verificar si se ejecuta como admin
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo.
    echo [ERROR] Este script debe ejecutarse como ADMINISTRADOR
    echo.
    echo Instrucciones:
    echo 1. Click derecho en este archivo
    echo 2. "Ejecutar como administrador"
    echo.
    pause
    exit /b 1
)

echo [OK] Ejecutando como administrador
echo.

REM Crear la tarea
echo [*] Creando tarea programada...
echo.

schtasks /create ^
    /tn "%TASK_NAME%" ^
    /tr "python \"%SCRIPT_PATH%\"" ^
    /sc daily ^
    /st %HORA% ^
    /f ^
    /ru SYSTEM

if %errorLevel% equ 0 (
    echo.
    echo [OK] Tarea creada exitosamente!
    echo.
    echo Detalles:
    echo   - Nombre: %TASK_NAME%
    echo   - Hora: %HORA% (09:00 AM)
    echo   - Frecuencia: Diaria
    echo   - Script: generar_reportes.py
    echo.
    
    REM Listar la tarea creada
    echo [*] Detalles de la tarea:
    schtasks /query /tn "%TASK_NAME%" /v
    
) else (
    echo.
    echo [ERROR] No se pudo crear la tarea
    echo Verifica que:
    echo   1. El script exista en: !SCRIPT_PATH!
    echo   2. Python este instalado y en el PATH
    echo   3. Ejecutes como ADMINISTRADOR
    echo.
)

echo.
echo ========================================
echo Para modificar o eliminar la tarea:
echo   Eliminar: schtasks /delete /tn "%TASK_NAME%" /f
echo   Modificar hora: schtasks /change /tn "%TASK_NAME%" /st HH:MM:SS
echo ========================================
echo.

pause
