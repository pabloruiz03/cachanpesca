@echo off
echo ===========================================
echo INSTALAR DEPENDENCIAS - CACHANPESCA
echo ===========================================
echo.

echo Instalando pywin32 (necesario para escuchar impresoras)...
pip install pywin32

echo.
echo Instalando otras dependencias...
pip install pypdf qrcode[pil] Pillow reportlab

echo.
echo ===========================================
echo INSTALACION COMPLETA
echo ===========================================
echo.
echo Ahora ejecuta:
echo   python source\capturar_etiquetas_admin.py
echo.
echo IMPORTANTE: Ejecuta como ADMINISTRADOR
echo.
pause
