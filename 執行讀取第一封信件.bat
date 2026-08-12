@echo off
chcp 65001 >nul
setlocal

cd /d "%~dp0"

echo ======================================
echo AutoEmail - 讀取最新一封信件
echo ======================================

where python >nul 2>nul
if %errorlevel% neq 0 (
  echo 找不到 Python，請先安裝 Python 3。
  echo 安裝後再重新執行此檔案。
  pause
  exit /b 1
)

python read_first_email.py

echo.
pause
