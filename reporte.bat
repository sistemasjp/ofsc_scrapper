@echo off
REM Genera el reporte de consumo de materiales para una fecha, sin
REM intervencion manual (usa las credenciales de .env para el login
REM automatico).
REM
REM Uso (desde cmd, o doble clic y luego escribe la fecha si te la pide):
REM   reporte.bat 2026-09-10

setlocal

if "%~1"=="" (
    echo Uso: reporte.bat AAAA-MM-DD
    exit /b 1
)

set "FECHA=%~1"
cd /d "%~dp0"

if not exist ".env" (
    echo No existe .env. Copia .env.example a .env y completa tus credenciales antes de correr esto.
    exit /b 1
)

if not exist "venv\Scripts\activate.bat" (
    echo No existe venv\. Sigue primero los pasos de instalacion del README.md.
    exit /b 1
)

call venv\Scripts\activate.bat
python ofsc_scraper.py --date %FECHA% --auto-login --headless --output "consumo_%FECHA%.xlsx"
