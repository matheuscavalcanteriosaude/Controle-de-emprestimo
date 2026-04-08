@echo
cd /d C:\Users\srsadmin\Desktop\notebooks - online 28-11
call py app.py
waitress-serve --host=0.0.0.0 --port=5000 app:app
