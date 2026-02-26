@echo off
title Actualizador
echo ===================================================
echo   Iniciando Script de Actualizacion en Segundo Plano
echo ===================================================

:: Cambiar al directorio donde se encuentra este archivo .bat
cd /d "%~dp0"

:: Ejecutar el script de Python
python Consulta_cotizaciones_auto.py

:: Pausar 5 segundos para que el usuario pueda leer el resultado si lo ejecuta manualmente
timeout /t 5