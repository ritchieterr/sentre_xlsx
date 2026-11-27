@echo off
title PDF to XLSX Extractor
echo.
echo ========================================
echo   PDF to XLSX Extractor
echo   Iniciando servidor...
echo ========================================
echo.
cd /d "%~dp0"
call .venv\Scripts\activate
python app.py
pause
