@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" py -3.11 -m venv .venv
call .venv\Scripts\activate.bat
pip install -r requirements.txt
python run.py
pause
