@echo off
title Compilar Reparador de Hipervinculos para Windows 7
echo.
echo IMPORTANTE: este proceso debe ejecutarse en Windows 7.
echo Se recomienda Python 3.8.10 de 64 bits (o 32 bits si Windows es de 32 bits).
echo.

py -3.8 -m pip install "pip<25"
py -3.8 -m pip install "openpyxl==3.1.5" "pyinstaller==5.13.2"

echo.
echo Compilando ejecutable...
py -3.8 -m PyInstaller --clean --onefile --windowed --name Reparador_Hipervinculos reparador_hipervinculos_gui.py

echo.
echo Si no hubo errores, el EXE esta en la carpeta dist.
pause
