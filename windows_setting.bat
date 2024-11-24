@echo off
REM Activate the virtual environment in the current directory
call windowsenv\Scripts\activate

REM 絕對指定虛擬環境內的 Python
windowsenv\Scripts\python.exe settingGUI.py .

REM 保持控制台開啟以查看執行結果
pause
