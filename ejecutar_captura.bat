@echo off
echo ===========================================
echo CAPTURAR ETIQUETAS - CACHANPESCA
echo ===========================================
echo.

REM Obtener directorio del script
set SCRIPT_DIR=%~dp0
set SCRIPT_DIR=%SCRIPT_DIR:~0,-1%

REM Verificar si ya estamos como administrador
net session >nul 2>&1
if %errorlevel% equ 0 (
    echo [OK] Ejecutando como administrador...
    echo.
    python "%SCRIPT_DIR%\source\capturar_etiquetas_admin.py"
) else (
    echo [INFO] Solicitando permisos de administrador...
    powershell -Command "Start-Process cmd -ArgumentList '/c cd /d %SCRIPT_DIR% && python source\capturar_etiquetas_admin.py' -Verb RunAs"
)

pause
