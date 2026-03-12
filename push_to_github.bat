@echo off
chcp 65001 >nul
set "GIT_PATH=C:\Program Files\Git\cmd"
if exist "%GIT_PATH%\git.exe" set "PATH=%GIT_PATH%;%PATH%"
cd /d "d:\CursorProject\AutomationAndAI"

git init
git add .
git commit -m "Add weekly_report_checker project"
git branch -M main
git remote add origin https://github.com/a5512561-creator/AutomationAndAI.git 2>nul
git push -u origin main

pause
