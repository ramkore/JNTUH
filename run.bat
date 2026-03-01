@echo off
echo ============================================
echo  B.Tech Internal Marks Automation System
echo ============================================
echo.

:: Check if Python is available
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed or not in PATH.
    echo Please install Python 3.8+ from https://www.python.org
    echo Make sure to check "Add Python to PATH" during installation.
    pause
    exit /b 1
)

:: Install requirements
echo Installing required packages...
pip install -r "%~dp0requirements.txt" --quiet
if %errorlevel% neq 0 (
    echo.
    echo WARNING: Some packages may have failed to install.
    echo Trying to continue anyway...
    echo.
)

:: Create desktop shortcut (if it doesn't exist yet)
set "SHORTCUT_NAME=Internal Marks App.lnk"

:: Try Public Desktop first (all users), fallback to current user Desktop
set "DESKTOP_PATH=%PUBLIC%\Desktop"
if not exist "%DESKTOP_PATH%" (
    set "DESKTOP_PATH=%USERPROFILE%\Desktop"
)

if not exist "%DESKTOP_PATH%\%SHORTCUT_NAME%" (
    echo Creating desktop shortcut...
    powershell -NoProfile -ExecutionPolicy Bypass -Command ^
        "$ws = New-Object -ComObject WScript.Shell; ^
         $s = $ws.CreateShortcut('%DESKTOP_PATH%\%SHORTCUT_NAME%'); ^
         $s.TargetPath = (Get-Command python -ErrorAction SilentlyContinue).Source; ^
         $s.Arguments = '\"%~dp0run.py\"'; ^
         $s.WorkingDirectory = '%~dp0'; ^
         $s.WindowStyle = 1; ^
         $s.Description = 'B.Tech Internal Marks Automation System'; ^
         $s.Save()"
    if %errorlevel% equ 0 (
        echo Desktop shortcut created: %DESKTOP_PATH%\%SHORTCUT_NAME%
    ) else (
        echo NOTE: Could not create shortcut on Public Desktop, trying user Desktop...
        set "DESKTOP_PATH=%USERPROFILE%\Desktop"
        powershell -NoProfile -ExecutionPolicy Bypass -Command ^
            "$ws = New-Object -ComObject WScript.Shell; ^
             $s = $ws.CreateShortcut('%DESKTOP_PATH%\%SHORTCUT_NAME%'); ^
             $s.TargetPath = (Get-Command python -ErrorAction SilentlyContinue).Source; ^
             $s.Arguments = '\"%~dp0run.py\"'; ^
             $s.WorkingDirectory = '%~dp0'; ^
             $s.WindowStyle = 1; ^
             $s.Description = 'B.Tech Internal Marks Automation System'; ^
             $s.Save()"
        if %errorlevel% equ 0 (
            echo Desktop shortcut created: %DESKTOP_PATH%\%SHORTCUT_NAME%
        ) else (
            echo WARNING: Could not create desktop shortcut. You can still run the app from run.bat.
        )
    )
) else (
    echo Desktop shortcut already exists.
)

echo.
echo Starting application...
echo.

:: Launch the app from project root
python "%~dp0run.py"

if %errorlevel% neq 0 (
    echo.
    echo ERROR: Application failed to start.
    echo Make sure all dependencies are installed:
    echo   pip install -r requirements.txt
    pause
)
