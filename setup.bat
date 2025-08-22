@echo off
cd /d "%~dp0MoonTranslator"

:: Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Python is not installed or not in PATH.
    echo Please install Python from https://www.python.org/downloads/
    echo After installation, make sure to check "Add Python to PATH" during installation.
    pause
    exit /b 1
)

:: Check if the virtual environment already exists
if exist "venv" (
    echo Virtual environment already exists. Activating it...
) else (
    :: Create a virtual environment
    python -m venv venv
    if %errorlevel% neq 0 (
        echo Failed to create virtual environment.
        exit /b 1
    )
)

:: Activate the virtual environment
call venv\Scripts\activate


pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
rmdir /s /q "dist"
pyinstaller MoonTranslator.py
copy *.ttf "dist\MoonTranslator"
copy config.json "dist\MoonTranslator"
echo Setup Complete!^(^>_^<^)! ...
pause
exit /b 1