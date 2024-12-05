@echo off
@echo *****************************************************************************
@echo ****************          check/ activate virtual environment    ************
@echo ****************          start the python script                ************
@echo *****************************************************************************


REM Ensure the current directory is set to the script's location
cd /d "%~dp0"

@echo off

REM Check if Python is installed
python --version >nul 2>&1
if %ERRORLEVEL% neq 0 (
   echo Python is not installed. Calling python_installer.bat...
   call "python_installer.bat"
   if %ERRORLEVEL% neq 0 (
       echo Failed to install Python. Please check your installation process.
       exit /b
   )
   echo Python has been successfully installed.
) else (
   echo Python is already installed.
)

REM Check if uv is installed by trying to get its version
uv --version >nul 2>&1
if %ERRORLEVEL% neq 0 (
   echo uv is not installed. Executing installation command...
   REM Use irm to download the script and execute it with iex
   powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
   REM Check if the installation command succeeded
   uv --version >nul 2>&1
   if %ERRORLEVEL% neq 0 (
       echo Failed to install uv. Please check your installation process.
       exit /b
   )
   echo uv has been successfully installed.
) else (
   echo uv is already installed.
)

REM Confirm uv version
echo Installed uv version:
uv --version

REM Check if the project is already initialized
if exist pyproject.toml (
   echo The project is already initialized.
) else (
   echo Initializing the project...
   uv init
   uv venv
   if %ERRORLEVEL% neq 0 (
       echo Failed to initialize the project. Please check for errors.
       exit /b
   )
   echo Project initialized successfully.
)

REM Define the requirements.txt path
set "requirements_file=requirements.txt"

REM Check if requirements.txt exists
if not exist "%requirements_file%" (
    echo requirements.txt not found.
    exit /b
)

REM Loop through each line in requirements.txt and add it using uv
for /f "delims=" %%d in (%requirements_file%) do (
    echo Adding dependency: %%d
    uv add %%d
)

echo All dependencies from requirements.txt have been added successfully.

uv sync

echo Starting monitoring
uv run main.py

REM Pause before exiting, allowing the user to view any output
pause
exit
