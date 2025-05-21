@echo off

:: === 參數設定 ===
set TASK_NAME=Babyview_Every5Min
set BAT_PATH=C:\Users\User\Babyview-Step1\run.bat
set FREQUENCY=MINUTE
set INTERVAL=5

:: === 建立排程任務 ===
schtasks /Create /F ^
 /SC %FREQUENCY% ^
 /MO %INTERVAL% ^
 /TN "%TASK_NAME%" ^
 /TR "\"%BAT_PATH%\""

echo.
echo ✅ 已建立每 %INTERVAL% 分鐘執行一次的排程任務：%TASK_NAME%
pause
