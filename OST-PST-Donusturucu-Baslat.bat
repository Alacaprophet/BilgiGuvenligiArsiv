@echo off
REM OST -> PST Donusturucu baslatici (Windows)
REM Python'un PATH'te oldugundan emin olun.
cd /d "%~dp0"
python run.py
if errorlevel 1 (
    echo.
    echo Program baslatilamadi. Python kurulu mu? "python --version" deneyin.
    pause
)
