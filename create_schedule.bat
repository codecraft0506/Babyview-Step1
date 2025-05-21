@echo off

:: === 設定參數 ===
set TASK_NAME=Babyview_GUI_5min
set BAT_PATH=C:\Users\User\Babyview-Step1\run.bat
set USERNAME=%USERNAME%
set FREQUENCY=MINUTE
set INTERVAL=5

:: === 建立排程任務，僅在使用者登入時執行，允許桌面互動 ===
schtasks /Create /F ^
 /SC %FREQUENCY% ^
 /MO %INTERVAL% ^
 /TN "%TASK_NAME%" ^
 /TR "\"%BAT_PATH%\"" ^
 /RL HIGHEST ^
 /RU "%USERNAME%" ^
 /IT

echo.
echo ✅ 已建立「可操作 GUI」的排程任務：%TASK_NAME%
pause
