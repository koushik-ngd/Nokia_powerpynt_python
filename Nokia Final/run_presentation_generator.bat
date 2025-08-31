@echo off
echo ========================================
echo Nokia Presentation Generator
echo ========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed or not in PATH
    echo.
    echo Please install Python from: https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation
    echo.
    pause
    exit /b 1
)

echo Python found! Checking dependencies...
echo.

REM Install required packages
echo Installing required Python packages...
python -m pip install --upgrade pip
python -m pip install python-pptx matplotlib numpy Pillow

if %errorlevel% neq 0 (
    echo.
    echo ERROR: Failed to install required packages
    echo Please check your internet connection and try again
    pause
    exit /b 1
)

echo.
echo Dependencies installed successfully!
echo.
echo Generating Nokia presentation...
echo.

REM Run the presentation generator
python nokia_presentation_generator.py

if %errorlevel% equ 0 (
    echo.
    echo ========================================
    echo SUCCESS! Presentation generated!
    echo ========================================
    echo.
    echo The file 'Nokia_Failure_Analysis_PowerPynt.pptx' has been created
    echo in the current directory.
    echo.
    echo You can now open it with Microsoft PowerPoint or any
    echo compatible presentation software.
    echo.
) else (
    echo.
    echo ERROR: Failed to generate presentation
    echo Please check the error messages above
    echo.
)

echo Press any key to exit...
pause >nul