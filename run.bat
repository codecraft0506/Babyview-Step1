@echo off
REM 切換到當前腳本所在資料夾
cd /d %~dp0

REM 啟用虛擬環境
call .venv\Scripts\activate.bat

REM 執行 Python 程式
python main.py

REM 結束後保持視窗開啟（可選）
pause
