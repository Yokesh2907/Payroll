@echo off
cd /d "%~dp0"
py -3 payroll_app.py 2>nul || python payroll_app.py
if errorlevel 1 pause
