@echo off
cd /d %~dp0

echo Dang khoi dong phan mem kiem tra xuat kho...
python -m streamlit run Stockchecker.py

pause
