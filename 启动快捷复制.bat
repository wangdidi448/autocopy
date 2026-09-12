@echo off
rem Launcher of Quick Workdesk (Qt). No console window kept.
set "SCRIPT=%~dp0qt_workdesk.py"

if exist "D:\miniconda\miniconda3\pythonw.exe" goto USE_MINICONDA
where pythonw >nul 2>nul && goto USE_PATH_PW
where py >nul 2>nul && goto USE_PY
goto NOPY

:USE_MINICONDA
start "" "D:\miniconda\miniconda3\pythonw.exe" "%SCRIPT%"
exit /b

:USE_PATH_PW
for /f "delims=" %%P in ('where pythonw') do (
    start "" "%%P" "%SCRIPT%"
    exit /b
)

:USE_PY
start "" py -3 "%SCRIPT%"
exit /b

:NOPY
mshta vbscript:msgbox("Python not found. Please install Python 3, then run: pip install PySide6",48,"Quick Workdesk")(window.close)
