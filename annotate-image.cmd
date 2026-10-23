@echo off
setlocal DisableDelayedExpansion
set "PYTHONUTF8=1"
if defined EASY_UIAUTO_IMAGE_PYTHON (
    "%EASY_UIAUTO_IMAGE_PYTHON%" "%~dp0scripts\annotate_image.py" %*
) else (
    python "%~dp0scripts\annotate_image.py" %*
)
exit /b %errorlevel%
