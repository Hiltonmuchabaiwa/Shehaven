@echo off
cd /d "%~dp0"
echo Installing requirements...
py -m pip install --upgrade -r requirements.txt
echo Starting SHEHaven Bedding...
py -m streamlit run app.py
pause
