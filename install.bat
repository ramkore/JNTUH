@echo off
setlocal EnableDelayedExpansion

title B.Tech Internal Marks - Installer
color 0A

echo.
echo  ============================================================
echo   B.Tech Internal Marks Automation System - Installer
echo  ============================================================
echo.
echo   This will set up everything you need to run the application.
echo.
echo  ============================================================
echo.

:: Track overall installation status
set "INSTALL_OK=1"

:: ---------------------------------------------------------------
::  STEP 1: Check if Python is installed
:: ---------------------------------------------------------------
echo [Step 1/5] Checking for Python...

set "PYTHON_CMD="

python --version >nul 2>&1
if %errorlevel% equ 0 ( set "PYTHON_CMD=python" & goto :python_found )

py --version >nul 2>&1
if %errorlevel% equ 0 ( set "PYTHON_CMD=py" & goto :python_found )

for %%P in (
    "%LOCALAPPDATA%\Programs\Python\Python314\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python310\python.exe"
    "C:\Python312\python.exe"
    "C:\Python311\python.exe"
) do (
    if exist %%P (
        set "PYTHON_CMD=%%~P"
        goto :python_found
    )
)

:: --- Python NOT found: Download and install ---
echo.
echo   Python is NOT installed. Downloading...
echo.

set "PYTHON_VERSION=3.12.2"
set "PYTHON_URL=https://www.python.org/ftp/python/%PYTHON_VERSION%/python-%PYTHON_VERSION%-amd64.exe"
set "INSTALLER_PATH=%TEMP%\python_installer.exe"

ping -n 1 www.python.org >nul 2>&1
if %errorlevel% neq 0 (
    echo   ERROR: No internet connection.
    echo   Install Python 3.8+ manually from https://www.python.org
    goto :install_failed
)

echo   Downloading Python %PYTHON_VERSION%...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; ^
     $ProgressPreference = 'SilentlyContinue'; ^
     Invoke-WebRequest -Uri '%PYTHON_URL%' -OutFile '%INSTALLER_PATH%' -UseBasicParsing"

if not exist "%INSTALLER_PATH%" (
    echo   ERROR: Download failed.
    goto :install_failed
)

echo   Installing Python silently (2-5 min)...
"%INSTALLER_PATH%" /quiet InstallAllUsers=0 PrependPath=1 Include_pip=1 Include_tcltk=1
del "%INSTALLER_PATH%" >nul 2>&1
set "PATH=%LOCALAPPDATA%\Programs\Python\Python312;%LOCALAPPDATA%\Programs\Python\Python312\Scripts;%PATH%"
timeout /t 3 >nul

python --version >nul 2>&1
if %errorlevel% equ 0 ( set "PYTHON_CMD=python" & goto :python_found )
py --version >nul 2>&1
if %errorlevel% equ 0 ( set "PYTHON_CMD=py" & goto :python_found )

echo   ERROR: Python install failed. Restart PC and try again.
goto :install_failed

:python_found
for /f "tokens=*" %%V in ('%PYTHON_CMD% --version 2^>^&1') do echo   Found: %%V
echo.

:: ---------------------------------------------------------------
::  STEP 2: Check packages and install missing
:: ---------------------------------------------------------------
echo [Step 2/5] Checking packages...
echo.

:: Delete temp file from any previous run
set "MISSING_FILE=%TEMP%\marks_missing_pkgs.txt"
if exist "%MISSING_FILE%" del "%MISSING_FILE%" >nul 2>&1

:: Run single check script — prints [SKIP] or [NEED] per package
%PYTHON_CMD% "%~dp0src\check_deps.py"

:: If the temp file exists, there are missing packages
if exist "%MISSING_FILE%" (
    set /p MISSING_PKGS=<"%MISSING_FILE%"
    del "%MISSING_FILE%" >nul 2>&1

    echo.
    echo   Installing: !MISSING_PKGS!
    echo   You will see download progress below:
    echo   ------------------------------------------
    %PYTHON_CMD% -m pip install !MISSING_PKGS!
    echo   ------------------------------------------

    if !errorlevel! equ 0 (
        echo   All packages installed successfully!
    ) else (
        echo   WARNING: Some packages may have failed.
        set "INSTALL_OK=0"
    )
) else (
    :: No missing file means all packages are present
    echo.
)
echo.

:: ---------------------------------------------------------------
::  STEP 3: Verify all packages
:: ---------------------------------------------------------------
echo [Step 3/5] Verifying dependencies...

%PYTHON_CMD% "%~dp0src\verify_deps.py"
if %errorlevel% neq 0 (
    echo   ERROR: One or more required packages failed verification.
    set "INSTALL_OK=0"
) else (
    echo   All dependencies verified successfully.
)
echo.

:: ---------------------------------------------------------------
::  STEP 4: Create desktop shortcut (All Users)
:: ---------------------------------------------------------------
echo [Step 4/5] Creating desktop shortcut...

set "SHORTCUT_NAME=Internal Marks App.lnk"
set "DESKTOP_PATH=%PUBLIC%\Desktop"
if not exist "%DESKTOP_PATH%" set "DESKTOP_PATH=%USERPROFILE%\Desktop"

:: Get the full Python executable path and derive pythonw.exe from it
for /f "tokens=*" %%A in ('%PYTHON_CMD% -c "import sys; print(sys.executable)"') do set "PYTHON_EXE=%%A"

:: Use pythonw.exe (no console window) — same directory as python.exe
set "PYTHONW_EXE=%PYTHON_EXE:python.exe=pythonw.exe%"
if not exist "%PYTHONW_EXE%" (
    echo   NOTE: pythonw.exe not found, using python.exe instead.
    set "PYTHONW_EXE=%PYTHON_EXE%"
)

:: Get the script directory (remove trailing backslash for clean paths)
set "SCRIPT_DIR=%~dp0"
if "%SCRIPT_DIR:~-1%"=="\" set "SCRIPT_DIR=%SCRIPT_DIR:~0,-1%"

:: Write a temporary PowerShell script to create the shortcut (avoids batch quoting issues)
set "PS_SCRIPT=%TEMP%\create_shortcut.ps1"

> "%PS_SCRIPT%" (
    echo $ws = New-Object -ComObject WScript.Shell
    echo $s = $ws.CreateShortcut^('%DESKTOP_PATH%\%SHORTCUT_NAME%'^)
    echo $s.TargetPath = '%PYTHONW_EXE%'
    echo $s.Arguments = '"%SCRIPT_DIR%\run.py"'
    echo $s.WorkingDirectory = '%SCRIPT_DIR%'
    echo $s.WindowStyle = 1
    echo $s.Description = 'B.Tech Internal Marks Automation System'
    echo $s.Save^(^)
    echo Write-Host 'Shortcut created successfully'
)

:: Execute the PowerShell script
powershell -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT%" 2>&1
set "SC_RESULT=%errorlevel%"
del "%PS_SCRIPT%" >nul 2>&1

if %SC_RESULT% equ 0 (
    if exist "%DESKTOP_PATH%\%SHORTCUT_NAME%" (
        echo   Desktop shortcut created successfully.
        goto :shortcut_done
    )
)

:: Fallback: try user-specific Desktop
echo   WARNING: Could not create on Public Desktop. Trying user Desktop...
set "DESKTOP_PATH=%USERPROFILE%\Desktop"

set "PS_SCRIPT=%TEMP%\create_shortcut.ps1"
> "%PS_SCRIPT%" (
    echo $ws = New-Object -ComObject WScript.Shell
    echo $s = $ws.CreateShortcut^('%DESKTOP_PATH%\%SHORTCUT_NAME%'^)
    echo $s.TargetPath = '%PYTHONW_EXE%'
    echo $s.Arguments = '"%SCRIPT_DIR%\run.py"'
    echo $s.WorkingDirectory = '%SCRIPT_DIR%'
    echo $s.WindowStyle = 1
    echo $s.Description = 'B.Tech Internal Marks Automation System'
    echo $s.Save^(^)
    echo Write-Host 'Shortcut created successfully'
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT%" 2>&1
set "SC_RESULT=%errorlevel%"
del "%PS_SCRIPT%" >nul 2>&1

if %SC_RESULT% equ 0 (
    if exist "%DESKTOP_PATH%\%SHORTCUT_NAME%" (
        echo   Desktop shortcut created on user Desktop.
    ) else (
        echo   ERROR: Shortcut creation failed.
        set "INSTALL_OK=0"
    )
) else (
    echo   ERROR: Shortcut creation failed.
    set "INSTALL_OK=0"
)

:shortcut_done
echo.

:: ---------------------------------------------------------------
::  STEP 5: Final Validation
:: ---------------------------------------------------------------
echo [Step 5/5] Running final validation...

:: Validate Python
%PYTHON_CMD% --version >nul 2>&1
if %errorlevel% neq 0 (
    echo   [FAIL] Python is not accessible.
    set "INSTALL_OK=0"
) else (
    echo   [OK] Python is installed and accessible.
)

:: Validate dependencies (re-check via import)
%PYTHON_CMD% -c "import pandas, numpy, openpyxl, reportlab" >nul 2>&1
if %errorlevel% neq 0 (
    echo   [FAIL] One or more Python packages are missing.
    set "INSTALL_OK=0"
) else (
    echo   [OK] All required packages are installed.
)

:: Validate shortcut
if exist "%DESKTOP_PATH%\%SHORTCUT_NAME%" (
    echo   [OK] Desktop shortcut exists at %DESKTOP_PATH%
) else (
    echo   [FAIL] Desktop shortcut was not created.
    set "INSTALL_OK=0"
)

echo.
echo  ============================================================
if "%INSTALL_OK%"=="1" (
    color 0A
    echo.
    echo   Installation Completed Successfully.
    echo.
    echo   You can now:
    echo     1. Double-click "Internal Marks App" on Desktop
    echo     2. Or double-click "run.bat" in this folder
    echo.
    echo  ============================================================
    echo.
    echo   Press any key to exit.
) else (
    color 0C
    echo.
    echo   Installation Failed. Check error messages above.
    echo.
    echo  ============================================================
    echo.
    echo   Press any key to exit.
)

goto :end

:install_failed
color 0C
echo.
echo  ============================================================
echo   INSTALLATION FAILED
echo   1. Check your internet connection
echo   2. Install Python 3.8+ from https://www.python.org
echo   3. Run install.bat again
echo  ============================================================
echo.
echo   Press any key to exit.

:end
echo.
pause
endlocal
