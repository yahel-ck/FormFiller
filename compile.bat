@echo off

@REM Save original working directory
PUSHD .

@REM Check if python is installed and alert user if not
py -3 -V > nul 2>&1
if errorlevel 1 (
    call :echodq "Python 3 is required but not found (Failed to run 'py' command)"
    goto :end
)

@REM Make sure python dependecies are installed
pyinstaller -v > nul 2>&1 || call :install_package pyinstaller || goto :end
call :ensure_package docxtpl || goto :end
call :ensure_package csv || goto :end
call :ensure_package winreg || goto :end

@REM Compile project into build folder
if not exist .\build mkdir .\build
CD .\build
if not exist ..\formfiller.py (
    call :echodq "Can't find 'formfiller.py' script to compile"
    goto :end
)
pyinstaller -F ..\formfiller.py
goto :end


@REM Functions
goto :eof

@REM Echos one parameter without quotation marks
:echodq
    echo %~1
goto :eof

@REM Installs a python package if the last command exited with error level 1
:install_package
    call :echodq "Installing package '%~1'"
    py -3 -m pip install %~1
    if errorlevel 1 call :echodq "Failed installing package '%~1'"
goto :eof

@REM Check if a package is installed and install it if not
:ensure_package
    py -3 -c "import %~1" > nul 2>&1 || call :install_package %~1
goto :eof

:end
    @REM Return to original working directory
    POPD
    @REM Pause if the script was ran without parameters
    if "%~1" == "" @pause
