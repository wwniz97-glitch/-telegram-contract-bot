@echo off
cd /d "%~dp0"
git status
set /p msg=Commit message: 
if "%msg%"=="" set msg=update
git add .
git commit -m "%msg%"
git push
pause
