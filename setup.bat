@echo off
setlocal

echo ================================================
echo CONFIGURADOR - CACHANPESCA
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

echo Python: %PYTHON%

REM Instalar dependencias
echo.
echo Instalando dependencias...
%PYTHON% -m pip install --upgrade pip --quiet
%PYTHON% -m pip install -r requirements.txt --quiet

if errorlevel 1 (
    echo ERROR: Fallo instalacion de dependencias
    pause
    exit /b 1
)

REM Crear carpetas
echo.
echo Creando estructura de carpetas...
if not exist etiquetas_json mkdir etiquetas_json
if not exist dist mkdir dist

REM Compilar QR Albaran
echo.
echo Compilando app_qr_albaran.exe...
if exist source (
    cd source
    pyinstaller app_qr_albaran.spec --onefile --noconfirm --clean 2>nul
    if exist dist\CACHANPESCA_QR_Albaran.exe (
        copy dist\CACHANPESCA_QR_Albaran.exe ..\dist\ 2>nul
        echo OK: dist\CACHANPESCA_QR_Albaran.exe creado
    )
    cd ..
)

REM Copiar config si existe
if exist source\bridge_config.json (
    copy source\bridge_config.json .\ 2>nul
)

REM Probar que funcionan las rutas de datos
echo.
echo Verificando rutas de datos...
if exist "C:\Users\Pablo\Downloads\rosi\HVETIQ CACHANPESCA\RES00" (
    echo OK: Datos encontrados en HVETIQ CACHANPESCA
)

echo.
echo ================================================
echo CONFIGURACION COMPLETADA
echo ================================================
echo.
echo Ejecutables en carpeta dist\
echo Etiquetas en etiquetas_json\
echo.
pause