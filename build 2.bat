@echo off
REM ============================================================
REM DET Lab Cover Sheet Tool - Windows Build Script
REM ============================================================
REM 
REM Prerequisites:
REM   1. Python 3.11+ installed
REM   2. Run this from the project directory
REM
REM Usage:
REM   build.bat          - Build the application
REM   build.bat clean    - Clean build artifacts and rebuild
REM ============================================================

echo ============================================
echo DET Lab Cover Sheet Tool - Build Script
echo ============================================
echo.

REM Check for clean flag
if "%1"=="clean" (
    echo Cleaning previous build artifacts...
    rmdir /s /q build 2>nul
    rmdir /s /q dist 2>nul
    del /q *.spec.bak 2>nul
    echo Clean complete.
    echo.
)

REM Create virtual environment if it doesn't exist
if not exist "venv" (
    echo Creating virtual environment...
    python -m venv venv
    echo Virtual environment created.
    echo.
)

REM Activate virtual environment
echo Activating virtual environment...
call venv\Scripts\activate.bat

REM Install/upgrade dependencies
echo Installing dependencies...
pip install --upgrade pip
pip install -r requirements.txt
pip install pyinstaller

echo.
echo Dependencies installed.
echo.

REM Build the executable
echo Building executable...
echo.
pyinstaller DETLAB.spec --noconfirm

echo.
if exist "dist\DET_Lab_CoverSheet\DET_Lab_CoverSheet.exe" (
    echo ============================================
    echo BUILD SUCCESSFUL!
    echo ============================================
    echo.
    echo Executable location:
    echo   dist\DET_Lab_CoverSheet\DET_Lab_CoverSheet.exe
    echo.
    echo To distribute:
    echo   Copy the entire 'dist\DET_Lab_CoverSheet' folder
    echo.
) else (
    echo ============================================
    echo BUILD FAILED
    echo ============================================
    echo Check the output above for errors.
)

pause
