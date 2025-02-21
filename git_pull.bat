@echo off
cd /d %~dp0

:: 丟棄所有本地更改並強制拉取遠端代碼
git reset --hard HEAD

:: 更改遠端 URL 為 HTTPS
git remote set-url origin https://github.com/Chrouos/MOST_committee.git

:: 拉取最新代碼
git pull

pause
