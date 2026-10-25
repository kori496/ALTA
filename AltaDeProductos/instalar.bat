@echo off
rem Instala Alta de productos: crea el acceso directo en el escritorio y en el menu Inicio.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0instalar.ps1"
echo.
pause
