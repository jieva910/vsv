@echo off

REM 获取当前BAT文件的目录路径（含结尾反斜杠）
set "RAW_PATH=%~dp0"
set "BAT_PATH=%RAW_PATH:~0,-1%"
REM 调用同目录下的PowerShell脚本，并传递路径参数
start /B powershell.exe  -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "%RAW_PATH%sap DI UI  connection for mass update.ps1" -Path "%BAT_PATH%"