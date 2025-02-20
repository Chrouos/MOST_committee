@echo off
cd /d %~dp0

:: 丟棄所有本地更改並強制拉取遠端代碼
git reset --hard HEAD
git pull

pause
