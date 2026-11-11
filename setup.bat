@echo off
setlocal
if not exist "%~dp0scripts\gui_setup.ps1" (
    echo Please extract the complete ZIP before running setup.bat.
    pause
    exit /b 1
)
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\gui_setup.ps1" %*
if errorlevel 1 (
    echo.
    echo GUI setup failed. See the message above, then run setup.bat again.
    pause
    exit /b 1
)
exit /b 0
