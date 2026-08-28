@echo off
REM ===========================================================================
REM  IntelliBI Operations Automation — full pipeline launcher (Task Scheduler)
REM  Point Task Scheduler here with "Start in" = this folder. No absolute paths.
REM ===========================================================================
cd /d "%~dp0"
if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" scripts\run_all.py
) else (
    python scripts\run_all.py
)
