@echo off
title Dong goi HT CREATIVE AGENT
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0dong_goi.ps1"
echo.
pause
