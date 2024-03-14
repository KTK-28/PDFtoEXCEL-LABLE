@echo off
echo Changing directory...
cd C:\Users\Owner\Desktop\PDFtoEXCEL-LABLE-main

echo Activating virtual environment...
call venv\Scripts\activate

echo Starting Flask server...
start "" python app.py

echo Flask server started.

echo Opening website...
timeout /t 5 /nobreak >nul
start http://localhost:5000