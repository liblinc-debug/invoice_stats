@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"
set "ROOT_DIR=%~dp0"

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

echo 正在准备 Windows 构建环境...
%PYTHON_CMD% -m venv ".build-venv"
if errorlevel 1 goto :build_failed
".build-venv\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 goto :build_failed
".build-venv\Scripts\python.exe" -m pip install -r "发票统计工具\requirements-build.txt"
if errorlevel 1 goto :build_failed

echo 正在构建 Windows 可执行程序...
set "PYINSTALLER_CONFIG_DIR=%ROOT_DIR%.build-venv\pyinstaller-config"
if not exist "%ROOT_DIR%build\icons" mkdir "%ROOT_DIR%build\icons"
"%ROOT_DIR%.build-venv\Scripts\python.exe" "%ROOT_DIR%发票统计工具\build_icon.py" windows "%ROOT_DIR%img\invoiceIcon.svg" "%ROOT_DIR%build\icons\invoiceIcon.ico"
if errorlevel 1 goto :build_failed
if not exist "%ROOT_DIR%build\icons\invoiceIcon.ico" goto :build_failed
"%ROOT_DIR%.build-venv\Scripts\python.exe" -m PyInstaller ^
    --noconfirm ^
    --clean ^
    --onefile ^
    --windowed ^
    --icon "%ROOT_DIR%build\icons\invoiceIcon.ico" ^
    --add-data "%ROOT_DIR%img\invoiceIcon.svg;img" ^
    --name "发票整理工具" ^
    --distpath "%ROOT_DIR%bin\windows" ^
    --workpath "%ROOT_DIR%build\pyinstaller-windows" ^
    --specpath "%ROOT_DIR%build" ^
    "%ROOT_DIR%发票统计工具\gui.py"
if errorlevel 1 goto :build_failed
if not exist "%ROOT_DIR%bin\windows\发票整理工具.exe" goto :build_failed
copy /y "%ROOT_DIR%LICENSE" "%ROOT_DIR%bin\windows\LICENSE" >nul
if errorlevel 1 goto :build_failed

echo.
echo 构建完成：%ROOT_DIR%bin\windows\发票整理工具.exe
pause
exit /b 0

:no_python
echo 未找到 Python，请先安装 Python 3.9 或更高版本。
pause
exit /b 1

:old_python
echo Python 版本过低，请安装 Python 3.9 或更高版本。
pause
exit /b 1

:build_failed
echo.
echo 构建失败，请检查网络和上方错误信息。
pause
exit /b 1
