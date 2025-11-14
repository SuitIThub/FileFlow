@echo off
cd /d "%~dp0"
python3 -m pip install -r requirements.txt
python3 copy_script.py
pause 