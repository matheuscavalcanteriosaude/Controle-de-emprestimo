@echo off
cd /d C:\Users\srsadmin\Desktop\notebooks - online 28-11
call venv\Scripts\activate.bat
waitress-serve --host=0.0.0.0 --port=5000 app:app
