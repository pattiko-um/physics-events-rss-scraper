@echo off
REM Build helper for Windows (run in project root from CMD)
python -m venv venv
call venv\Scripts\activate
python -m pip install --upgrade pip
pip install -r requirements.txt pyinstaller
pyinstaller --noconfirm --onefile --name physics-events-scraper main.py
echo Built: dist\physics-events-scraper.exe
