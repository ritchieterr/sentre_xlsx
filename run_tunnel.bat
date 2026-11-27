@echo off
title Tunel Cloudflare - Acceso Remoto
echo Iniciando tunel de Cloudflare...
echo Tu URL publica aparecera abajo como: https://xxxx-xxxx.trycloudflare.com
echo.
"%LOCALAPPDATA%\Microsoft\WinGet\Packages\Cloudflare.cloudflared_Microsoft.Winget.Source_8wekyb3d8bbwe\cloudflared.exe" tunnel --url http://localhost:5000
pause
