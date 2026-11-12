@echo off
chcp 65001 >nul
title Tai video tu Excel
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
  echo [!] Chua cai Python. Tai tai https://www.python.org/downloads/
  echo     Khi cai nho tick o "Add python.exe to PATH".
  pause
  exit /b 1
)

echo Dang kiem tra / cap nhat thu vien (yt-dlp, openpyxl, ffmpeg)...
python -m pip install -q -U --disable-pip-version-check yt-dlp openpyxl imageio-ffmpeg
python -m pip install -q -U --disable-pip-version-check "yt-dlp[default,curl-cffi]"
if errorlevel 1 echo [!] Khong cai duoc curl_cffi - trang dung Cloudflare co the loi.
python -m pip install -q -U --disable-pip-version-check playwright
if errorlevel 1 echo [!] Khong cai duoc playwright - xem thong bao khi tool chay.
echo.
python tai_video.py %*
echo.
pause
