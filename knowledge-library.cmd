@echo off
setlocal
call "%~dp0setup.bat" %*
exit /b %errorlevel%
