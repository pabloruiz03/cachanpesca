@echo off
echo Ejecutando lector de spooler como ADMINISTRADOR...
powershell -Command "Start-Process python -ArgumentList 'c:\Users\Pablo\Downloads\rosi\HVETIQ CACHANPESCA\leer_spooler_admin.py' -Verb RunAs"
pause