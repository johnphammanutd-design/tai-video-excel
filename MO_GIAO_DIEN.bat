@echo off
chcp 65001 >nul
title HT CREATIVE AGENT - Giao dien
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
  echo [!] Chua cai Python. Tai tai https://www.python.org/downloads/
  echo     Khi cai nho tick o "Add python.exe to PATH".
  pause
  exit /b 1
)

echo Dang cap nhat bo tai video yt-dlp...
python -m pip install -q -U --disable-pip-version-check yt-dlp openpyxl imageio-ffmpeg
echo Dang mo giao dien... (de cua so nay mo trong luc dung)
python giao_dien.py %*
if errorlevel 1 pause
