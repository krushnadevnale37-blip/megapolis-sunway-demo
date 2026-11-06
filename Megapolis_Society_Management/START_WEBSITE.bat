@echo off
title Megapolis Defaulter Report Website
cd /d "%~dp0"
echo Installing required Python packages...
py -m pip install -r requirements.txt
if errorlevel 1 (
  echo.
  echo Package installation failed. Please check Python installation and internet.
  pause
  exit /b 1
)
echo.
echo Starting your local website...
echo Keep this black window open while using the website.
start "" http://127.0.0.1:5000
py app.py
pause
