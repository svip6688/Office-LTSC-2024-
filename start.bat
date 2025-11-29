@echo off
REM 直接启动 PowerShell Office 安装脚本，不闪退
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0OfficeInstaller.ps1"
pause
