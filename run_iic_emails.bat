@echo off
REM Activate the virtual environment
call "C:\Users\johnv\IIC-email-scheduler\venv\Scripts\activate.bat"

REM Run the Python script and log output
python "C:\Users\johnv\IIC-email-scheduler\livev1.py" > "C:\Users\johnv\IIC-email-scheduler\log.txt" 2>&1

REM Deactivate the virtual environment (optional, Task Scheduler will close the process)
deactivate
