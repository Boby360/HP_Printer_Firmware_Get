@echo off
setlocal enabledelayedexpansion
echo This script temporarily bypasses the powershell executionpolicy for the HP Firmware Get script.
echo:
REM Get the full path of the batch file
REM set "batchfile=%~f0"

REM Extract the directory from the batch file path
set "batchdir=%~dp0"


Powershell.exe -executionpolicy bypass -Command %batchdir%HP_Firmware_get_dev_autoread.ps1

if %errorlevel% neq 0 (
    echo Failed
	REM checking if script ran as Administrator    
    net session >nul 2>&1
    if !errorLevel! == 0 (
	echo 
    ) else (
        echo It seems like you might need to run this script as an Administrator.
		echo Right click this file, and select "Run as administrator" if possible.
		pause
		exit
    )
) else (
    REM echo PowerShell command succeeded!
    REM Do something here if the command succeeds
)

pause