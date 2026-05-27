@echo off
echo ========================================
echo   JinhuaJuhuo Table Processor - Build
echo ========================================

echo Checking Python...
python --version
if errorlevel 1 (
    echo [ERROR] Python not found.
    pause
    exit /b 1
)

echo [1/3] Installing dependencies...
pip install pandas==3.0.1 "openpyxl>=3.1.5" xlrd==2.0.2 numpy==2.4.2 pyinstaller --force-reinstall openpyxl
if errorlevel 1 (
    echo [ERROR] Failed to install.
    pause
    exit /b 1
)

echo [2/3] Cleaning old build...
rmdir /s /q build dist 2>nul
del /q *.spec 2>nul
echo Building EXE...
python -m PyInstaller --noconsole --onedir --name JinhuaJuhuo --icon favicon.ico "金华聚火表格处理_整合版_fixed.py"
if errorlevel 1 (
    echo [ERROR] Build failed!
    pause
    exit /b 1
)

echo [3/3] Done.
echo Output: dist\JinhuaJuhuo\JinhuaJuhuo.exe
pause
start dist\JinhuaJuhuo
