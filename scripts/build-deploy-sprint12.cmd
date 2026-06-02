@echo off
setlocal

@echo Build project and deploy to sprint12.sitefpo.com

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "& '%~dp0build.ps1' -RemoteDeployTarget @{Url='https://sprint12.sitefpo.com/installCollection'; SiteName='sprint12'}"
set deployExit=%errorlevel%

pause
exit /b %deployExit%
