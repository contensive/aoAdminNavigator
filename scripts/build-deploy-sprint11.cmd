@echo off
setlocal

@echo Build project and deploy to sprint11.sitefpo.com

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "& '%~dp0build.ps1' -RemoteDeployTarget @{Url='https://sprint11.sitefpo.com/installCollection'; SiteName='sprint11'}"
set deployExit=%errorlevel%

pause
exit /b %deployExit%
