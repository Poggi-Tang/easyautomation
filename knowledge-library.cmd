@echo off
setlocal
cd /d "%~dp0"
set "PYTHONPATH=%~dp0src;%PYTHONPATH%"
if exist ".venv\Scripts\pythonw.exe" (
    start "" ".venv\Scripts\pythonw.exe" -m easy_uiauto.library_launcher %*
) else (
    start "" pythonw.exe -m easy_uiauto.library_launcher %*
)
