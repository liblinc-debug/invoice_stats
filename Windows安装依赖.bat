@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

where py >nul 2>nul
if not errorlevel 1 (
    set "PYTHON_CMD=py -3"
) else (
    where python >nul 2>nul
    if errorlevel 1 goto :no_python
    set "PYTHON_CMD=python"
)

%PYTHON_CMD% -c "import sys; raise SystemExit(0 if sys.version_info >= (3, 9) else 1)"
if errorlevel 1 goto :old_python

echo 正在创建 Python 虚拟环境...
%PYTHON_CMD% -m venv ".venv"
if errorlevel 1 goto :install_failed

echo 正在安装依赖...
".venv\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 goto :install_failed
".venv\Scripts\python.exe" -m pip install -r "发票统计工具\requirements.txt"
if errorlevel 1 goto :install_failed

echo.
echo 依赖安装完成。以后双击“Windows启动统计.bat”即可。
pause
exit /b 0

:no_python
echo 未找到 Python。请先从 https://www.python.org/downloads/windows/ 安装 Python 3.9 或更高版本。
echo 安装时请勾选“Add Python to PATH”。
pause
exit /b 1

:old_python
echo Python 版本过低，请安装 Python 3.9 或更高版本。
pause
exit /b 1

:install_failed
echo.
echo 依赖安装失败，请检查网络连接和上方错误信息。
pause
exit /b 1
