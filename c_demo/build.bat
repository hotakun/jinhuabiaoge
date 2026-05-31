@echo off
echo Building C demo...
gcc -O2 -Wall -o c_demo.exe main.c
if errorlevel 1 (
    echo Build failed. Install gcc first.
    pause
    exit /b 1
)
echo Done.
pause
