@echo off
cd /d "%~dp0"
echo 正在啟動 蒐證照片報表產生器...
echo Starting Photo Report Generator...
echo.
streamlit run src/app.py
pause
