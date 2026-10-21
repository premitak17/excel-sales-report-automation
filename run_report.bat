@echo off
setlocal
cd /d "%~dp0"
python -m sales_report %*
if errorlevel 1 exit /b %errorlevel%
endlocal