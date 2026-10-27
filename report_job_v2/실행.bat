@echo off
setlocal
cd /d "%~dp0"
where py >nul 2>nul
if errorlevel 1 goto use_python
py -3 report_job_v2_single.py
goto finish
:use_python
python report_job_v2_single.py
:finish
if errorlevel 1 pause
endlocal
