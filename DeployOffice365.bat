@echo off
SETLOCAL

echo ================================
echo Installing Microsoft 365 Apps 64-bit...
echo ================================

REM Set temp folder for download
SET TempODT=%TEMP%\ODT
IF NOT EXIST "%TempODT%" mkdir "%TempODT%"

REM Download setup.exe from GitHub
echo Downloading ODT...
powershell -Command "Invoke-WebRequest -Uri 'https://github.com/MrWes3000/SPARK/raw/refs/heads/main/setup.exe' -OutFile '%TempODT%\setup.exe'"

REM Run setup.exe with your GitHub XML
echo Installing Office...
"%TempODT%\setup.exe" /configure "https://raw.githubusercontent.com/MrWes3000/SPARK/95a0d473c7c40c854df33bf37cc48d3b9e1a08f1/Office64_Full_NoGroove.xml"

REM Optional: clean up
echo Cleaning up temporary files...
rmdir /s /q "%TempODT%"

echo ================================
echo Installation complete. Please have the user sign in to activate Office.
echo ================================
pause
ENDLOCAL
