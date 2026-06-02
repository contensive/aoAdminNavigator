@echo off
setlocal

@echo Build project and install on site: veronica

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "& '%~dp0build.ps1' -LocalDeployTarget 'veronica'"
set deployExit=%errorlevel%

pause
exit /b %deployExit%
