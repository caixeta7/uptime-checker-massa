@echo off
TITLE SystemInfo Uptime Runner
color 0B
echo ========================================================
echo  Iniciando SystemInfo Uptime (Modo Otimizado)
echo ========================================================
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Run-Uptime.ps1"
pause
