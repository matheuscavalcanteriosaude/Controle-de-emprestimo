@echo
cd /d D:\FSTEC\notebooks - online 28-11
call venv\scripts\acivate.bat
waitress-serve --host=0.0.0.0 --port=5000 app:app
