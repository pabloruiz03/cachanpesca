@echo off
setlocal

:: CACHANPESCA keep_alive - Indetectable auto-restart service
:: Se ejecuta siempre, aunque se elimine del Task Manager
:: Se regenera a los 3 segundos si muere

set "SCRIPT_DIR=%~dp0"
set "PS1_FILE=%SCRIPT_DIR%keep_alive_service.ps1"
set "BAT_FILE=%SCRIPT_DIR%keep_alive.bat"

:: Crear script PowerShell si no existe
if not exist "%PS1_FILE%" (
    powershell -ExecutionPolicy Bypass -Command ^
        "Set-Content -Path '%PS1_FILE%' -Value 'while($true){if(!(Get-Process -Name CapturarEtiquetas -ErrorAction SilentlyContinue)){Start-Process \"%SCRIPT_DIR%CapturarEtiquetas.exe\" -WindowStyle Hidden};Start-Sleep -Seconds 3}' -Encoding UTF8"
)

:: Buscar si ya hay uno corriendo
wmic process where "name='wscript.exe' or name='cscript.exe' or name='powershell.exe'" get commandline 2>nul | findstr /i "keep_alive" >nul
if %errorlevel%==0 goto :already_running

:: Lanzar de forma oculta (sin ventana, sin interfaz)
cscript //nologo "%SCRIPT_DIR%\..\launchers\start_hidden.vbs" "%PS1_FILE%"

:already_running
exit /b 0