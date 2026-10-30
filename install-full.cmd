@echo off
setlocal
set "EASY_UIAUTO_BOOTSTRAP_FILE=%~dp0scripts\full_setup.ps1"
if exist "%EASY_UIAUTO_BOOTSTRAP_FILE%" goto run
set "EASY_UIAUTO_BOOTSTRAP_FILE=%TEMP%\easy-uiauto-setup-%RANDOM%-%RANDOM%.ps1"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='Stop'; Invoke-WebRequest -UseBasicParsing -Uri 'https://github.com/Poggi-Tang/easyautomation/releases/latest/download/easy-uiauto-setup.ps1' -OutFile $env:EASY_UIAUTO_BOOTSTRAP_FILE"
if errorlevel 1 exit /b 1
:run
powershell -NoProfile -ExecutionPolicy Bypass -File "%EASY_UIAUTO_BOOTSTRAP_FILE%" -Action install %*
exit /b %errorlevel%
