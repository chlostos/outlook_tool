@echo off
@echo *****************************************************************************
@echo ****************          Download and Install Python          **************
@echo *****************************************************************************

REM Define the download URL and target location
set "python_url=https://www.python.org/ftp/python/3.11.0/python-3.11.0-amd64.exe"
set "installer_path=c:\veera\python-3.11.0-amd64.exe"

REM Ensure the target directory exists
if not exist "c:\veera" mkdir "c:\veera"

REM Download Python installer using PowerShell
echo Downloading Python 3.11...
powershell -Command "Invoke-WebRequest -Uri '%python_url%' -OutFile '%installer_path%' -UseBasicParsing"
if not exist "%installer_path%" (
    echo Failed to download Python installer.
    exit /b 1
)

echo Download complete.

REM Install Python silently
echo Installing Python 3.11...
"%installer_path%" /quiet InstallAllUsers=1 PrependPath=1 Include_test=0
if %ERRORLEVEL% neq 0 (
    echo Python installation failed. Please check for errors.
    exit /b 1
)

echo Installation complete.

REM Verify Python installation
python --version >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo Python is not accessible after installation. Adding it to PATH...
    REM Add Python to PATH manually
    setx /M path "%path%;C:\Program Files\Python311\"
    if %ERRORLEVEL% neq 0 (
        echo Failed to modify PATH. Please check your permissions.
        exit /b 1
    )
)

echo Python installed and PATH updated successfully.

exit /b 0
