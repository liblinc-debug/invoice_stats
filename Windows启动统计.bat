@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" goto :not_installed

".venv\Scripts\python.exe" "发票统计工具\gui.py"
set "EXIT_CODE=%ERRORLEVEL%"
if not "%EXIT_CODE%"=="0" (
    echo.
    echo 程序异常退出，请根据上方提示检查。
    pause
)
exit /b %EXIT_CODE%

:not_installed
echo 尚未安装依赖，请先双击“Windows安装依赖.bat”。
pause
exit /b 1
