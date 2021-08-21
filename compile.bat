@echo off

@REM Check if python is installed and alert user if not
py -3 -V > nul 2>&1
if errorlevel 1 (
    echo Please install Python 3
    goto :end
)


@REM Make sure python dependecies are installed
pyinstaller --help > nul 2>&1
call :install_if_last_command_failed pyinstaller || goto :end
call :require_package docxtpl || goto :end
call :require_package csv || goto :end
call :require_package winreg || goto :end


@REM Compile project into build folder
if not exist .\build mkdir .\build
CD .\build
pyinstaller -F ..\formfiller.py
goto :end


@REM Functions

@REM Installs a python package if the last command exited with error level 1
:install_if_last_command_failed
    if errorlevel 1 (
        echo Package '%~1' not found, installing now
        py -3 -m pip install %~1
        exit /b 1
    )
    exit /b 0

@REM Check if a package is installed and install it if not
:require_package
    py -3 -c "import %~1" > nul 2>&1
    call :install_if_last_command_failed %~1
    exit /b %errorlevel%

:end
    if "%~1" == "" @pause
