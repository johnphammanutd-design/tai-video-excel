@echo off
title Dong goi Tai Video Excel
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0dong_goi.ps1"
echo.
pause
