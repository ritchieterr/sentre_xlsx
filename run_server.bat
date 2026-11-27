@echo off
title Servidor - Sistema de Inventarios OPB
echo Iniciando el servidor...
cd /d "%~dp0"
call .venv\Scripts\activate
python -c "from waitress import serve; from app import app; print('Servidor corriendo en http://localhost:5000'); print('Para detener, cierra esta ventana'); serve(app, host='0.0.0.0', port=5000)"
pause
