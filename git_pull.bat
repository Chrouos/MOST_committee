@echo off
cd /d %~dp0

:: Drop the all changes in the working directory
git reset --hard HEAD

:: Delete all untrack data and folder
git clean -fd

:: Change the remote URL to HTTPS
git remote set-url origin https://github.com/Chrouos/MOST_committee.git

:: Pull the latest changes from the remote repository
git pull

pause
