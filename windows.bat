@echo off
@chcp 65001 >nul

REM 確保虛擬環境已存在
if not exist "venv\Scripts\activate" (
    echo [錯誤] 虛擬環境未創建，請先執行 setup_env.bat 來設置環境。
    pause
    exit /b 1
)

REM 啟動虛擬環境
call .\venv\Scripts\activate

REM 顯式使用虛擬環境的 Python
echo 正在執行 mainGUI.py 腳本...
.\venv\Scripts\python.exe mainGUI.py .

REM 保留控制台開啟，顯示執行結果
pause
