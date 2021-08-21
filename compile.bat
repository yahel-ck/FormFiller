@echo off

REM Check if pyinstaller is installed and install if not
pyinstaller --help > nul 2>&1
if errorlevel 1 (
    echo Package 'pyinstaller' is not installed, installing now
    pip install pyinstaller
)

if not exist .\build mkdir .\build
CD .\build
pyinstaller -F ..\formfiller.py
@pause
