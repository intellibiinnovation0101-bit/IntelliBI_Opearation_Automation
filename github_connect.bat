@echo off
setlocal
cd /d "%~dp0"
echo ============================================================
echo   Connect IntelliBI_Operations_Automation to GitHub
echo   repo: https://github.com/intellibiinnovation0101-bit/IntelliBI_Opearation_Automation
echo   (note the repo spelling: "Opearation")
echo ============================================================
echo.

REM --- 0. clear any stale git lock ---
if exist ".git\index.lock" del /f /q ".git\index.lock"

REM --- 0b. make sure a commit identity exists ---
set "GITEMAIL="
for /f "delims=" %%i in ('git config user.email 2^>nul') do set "GITEMAIL=%%i"
if not defined GITEMAIL call :setid

REM --- 1. initialise git if needed, use 'main' ---
if not exist ".git" (
    git init
)
git branch -M main

REM --- 2. stage everything (secrets are excluded by .gitignore) ---
git add .

echo ============================================================
echo   REVIEW the files that will be committed below.
echo   Confirm NONE of these appear:
echo     credentials\  .venv\  logs\  cache\  temp\  data_inputs\  output\
echo   (only .example / README files under credentials\ and data_inputs\ are OK)
echo ============================================================
git status
echo.
echo   If the list looks correct, continue. Otherwise close this window.
pause

git commit -m "Initial commit: IntelliBI Operations Automation"

REM --- 3. point 'origin' at the CORRECT Operations repo ---
git remote remove origin 2>nul
git remote add origin https://github.com/intellibiinnovation0101-bit/IntelliBI_Opearation_Automation.git
echo --- remote (MUST say IntelliBI_Opearation_Automation, NOT the Sales repo) ---
git remote -v
echo.
echo   Confirm the remote above is the OPERATIONS repo, then continue.
pause

REM --- 4. publish to main (force replaces GitHub's placeholder README only) ---
git push -u origin main --force
if errorlevel 1 goto :err

REM --- 5. create dev and prod from main and publish them ---
git branch dev 2>nul
git branch prod 2>nul
git push -u origin dev
git push -u origin prod

REM --- 6. verify ---
echo.
echo --- verification: branches (local + remote) ---
git fetch --all --prune
git branch -a
echo --- upstream tracking ---
git branch -vv
echo.
echo DONE.  main / dev / prod are on GitHub (IntelliBI_Opearation_Automation).
goto :end

:setid
echo No git identity found. Enter it once (used on your commits):
set /p GNAME=Your name :
set /p GEMAIL=Your email:
git config --global user.name "%GNAME%"
git config --global user.email "%GEMAIL%"
goto :eof

:err
echo.
echo *** PUSH FAILED ***
echo Check that PyCharm/Git is logged in to GitHub as an account with
echo WRITE access to intellibiinnovation0101-bit/IntelliBI_Opearation_Automation.
echo (GitHub needs OAuth login or a Personal Access Token - not a password.)

:end
echo.
pause
endlocal
