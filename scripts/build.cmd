@echo off
setlocal

@echo Build aoAdminNavigator

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0build.ps1"

if not "%1"=="/nopause" pause
