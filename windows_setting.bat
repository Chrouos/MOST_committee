@echo off
REM Activate the virtual environment in the current directory
call windowsenv\Scripts\activate

REM Execute the settingGUI.py script in the current directory, passing the current directory as the repository path
python settingGUI.py .

REM Pause to keep the console open and display the result
pause
